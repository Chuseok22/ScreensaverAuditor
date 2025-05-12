// Services/ScreensaverAuditorService.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Xml;
using System.Diagnostics.Eventing.Reader;
using ScreensaverAuditor.Models;

namespace ScreensaverAuditor.Services
{
    public class ScreensaverAuditorService
    {
        private const string AuditSubcategory = "기타 로그온/로그오프 이벤트";

        public void EnableAuditPolicy()
        {
            if (!IsAdministrator())
                throw new UnauthorizedAccessException("관리자 권한이 필요합니다.");

            // SET
            RunAuditPol($"/set /subcategory:\"{AuditSubcategory}\" /success:enable /failure:enable", "감사 정책 설정 실패");
            // GET 확인
            RunAuditPol($"/get /subcategory:\"{AuditSubcategory}\"", outputOnly: true);
        }

        private static void RunAuditPol(string args, string? errorMsg = null, bool outputOnly = false)
        {
            var psi = new ProcessStartInfo("auditpol.exe", args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = outputOnly
            };
            using var proc = Process.Start(psi) ?? throw new Exception("auditpol 실행 실패");
            string output = outputOnly ? proc.StandardOutput.ReadToEnd() : string.Empty;
            proc.WaitForExit();
            if (proc.ExitCode != 0 && errorMsg != null) throw new Exception(errorMsg);
            if (outputOnly) Console.WriteLine(output);
        }

        public bool IsAuditPolicyEnabled()
        {
            try
            {
                var processStartInfo = new ProcessStartInfo("auditpol.exe", $"/get /subcategory:\"{AuditSubcategory}\"")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using var process = Process.Start(processStartInfo) ?? throw new Exception("auditpol 실행 실패");
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                // 출력에서 "성공 및" 또는 "Success and" 문자열이 포함되어 있는지 확인
                return output.Contains("성공 및") || output.Contains("Success and");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"감사 정책 확인 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public List<ScreensaverEvent> GetScreensaverEvents(DateTime from, DateTime to, string? user = null)
        {
            var events = new List<ScreensaverEvent>();
            Console.WriteLine($"조회 시작: {from:yyyy-MM-dd HH:mm:ss} ~ {to:yyyy-MM-dd HH:mm:ss}");

            try
            {
                // PowerShell 명령어 구성
                string command = $@"
            Get-WinEvent -FilterHashtable @{{
                LogName='Security'
                ID=4802,4803
                StartTime=[datetime]'{from:yyyy-MM-dd HH:mm:ss}'
                EndTime=[datetime]'{to:yyyy-MM-dd HH:mm:ss}'
            }}";

                if (user != null)
                {
                    // 사용자 이름에 알파벳, 숫자, 공백, 점, 대시, 밑줄만 허용
                    if (!System.Text.RegularExpressions.Regex.IsMatch(user, @"^[a-zA-Z0-9\s\._\-]+$"))
                        throw new ArgumentException("사용자 이름에 허용되지 않는 문자가 포함되어 있습니다.");
                    command += $" | Where-Object {{ $_.Message -like '*{user}*' }}";
                }

                command += " | Format-List TimeCreated,Id,MachineName,Message";

                // PowerShell 실행
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{command}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                Console.WriteLine("PowerShell 명령 실행: " + command);

                using (var process = Process.Start(startInfo))
                {
                    if (process == null) throw new Exception("PowerShell 실행 실패");

                    // 출력 읽기
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        Console.WriteLine("PowerShell 오류:");
                        Console.WriteLine(error);
                        throw new Exception($"PowerShell 오류: {error}");
                    }

                    // 이벤트 파싱
                    var currentEvent = new Dictionary<string, string>();
                    var messageLines = new List<string>();
                    bool inMessage = false;

                    foreach (var line in output.Split('\n'))
                    {
                        string trimmedLine = line.Trim();

                        if (string.IsNullOrWhiteSpace(trimmedLine))
                        {
                            if (currentEvent.Count > 0)
                            {
                                // 메시지 라인 병합
                                if (messageLines.Count > 0)
                                {
                                    currentEvent["Message"] = string.Join("\n", messageLines);
                                }
                                try
                                {
                                    if (currentEvent.TryGetValue("TimeCreated", out var timeCreatedStr) &&
                                        currentEvent.TryGetValue("Id", out var idStr) &&
                                        DateTime.TryParse(timeCreatedStr, out var timeCreated) &&
                                        int.TryParse(idStr, out var id))
                                    {
                                        string computerName = currentEvent.TryGetValue("MachineName", out var comp) ? comp : "Unknown";
                                        string message = currentEvent.TryGetValue("Message", out var msg) ? msg : "";

                                        // 메시지에서 정보 추출
                                        string username = "Unknown";
                                        string domain = "Unknown";
                                        string securityId = "Unknown";
                                        string logonId = "Unknown";
                                        string sessionId = "Unknown";

                                        // 메시지 파싱
                                        foreach (var msgLine in message.Split('\n'))
                                        {
                                            string msgTrimmed = msgLine.Trim();

                                            if (msgTrimmed.StartsWith("계정 이름:") || msgTrimmed.StartsWith("Account Name:"))
                                            {
                                                var parts = msgTrimmed.Split(':', 2);
                                                if (parts.Length > 1)
                                                {
                                                    var accountParts = parts[1].Trim().Split('\\');
                                                    if (accountParts.Length > 1)
                                                    {
                                                        domain = accountParts[0].Trim();
                                                        username = accountParts[1].Trim();
                                                    }
                                                    else
                                                    {
                                                        username = parts[1].Trim();
                                                    }
                                                }
                                            }
                                            else if (msgTrimmed.StartsWith("계정 도메인:") || msgTrimmed.StartsWith("Account Domain:"))
                                            {
                                                var parts = msgTrimmed.Split(':', 2);
                                                if (parts.Length > 1)
                                                {
                                                    domain = parts[1].Trim();
                                                }
                                            }
                                            else if (msgTrimmed.StartsWith("보안 ID:") || msgTrimmed.StartsWith("Security ID:"))
                                            {
                                                var parts = msgTrimmed.Split(':', 2);
                                                if (parts.Length > 1)
                                                {
                                                    securityId = parts[1].Trim();
                                                }
                                            }
                                            else if (msgTrimmed.StartsWith("로그온 ID:") || msgTrimmed.StartsWith("Logon ID:"))
                                            {
                                                var parts = msgTrimmed.Split(':', 2);
                                                if (parts.Length > 1)
                                                {
                                                    logonId = parts[1].Trim();
                                                }
                                            }
                                            else if (msgTrimmed.StartsWith("세션 ID:") || msgTrimmed.StartsWith("Session ID:"))
                                            {
                                                var parts = msgTrimmed.Split(':', 2);
                                                if (parts.Length > 1)
                                                {
                                                    sessionId = parts[1].Trim();
                                                }
                                            }
                                        }

                                        // 이벤트 추가
                                        var evt = new ScreensaverEvent(
                                            timeCreated,
                                            id,
                                            computerName,
                                            message,
                                            username,
                                            domain,
                                            securityId,
                                            logonId,
                                            sessionId);

                                        events.Add(evt);
                                        Console.WriteLine($"이벤트 추가: ID={id}, 시간={timeCreated:yyyy-MM-dd HH:mm:ss}, 사용자={username}, PC={computerName}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"이벤트 처리 오류: {ex.Message}");
                                }

                                currentEvent.Clear();
                                messageLines.Clear();
                                inMessage = false;
                            }
                            continue;
                        }

                        if (trimmedLine.StartsWith("TimeCreated") ||
                            trimmedLine.StartsWith("Id") ||
                            trimmedLine.StartsWith("MachineName"))
                        {
                            int colonPos = trimmedLine.IndexOf(':');
                            if (colonPos > 0)
                            {
                                string key = trimmedLine.Substring(0, colonPos).Trim();
                                string value = trimmedLine.Substring(colonPos + 1).Trim();
                                currentEvent[key] = value;
                            }
                        }
                        else if (trimmedLine.StartsWith("Message"))
                        {
                            inMessage = true;
                            int colonPos = trimmedLine.IndexOf(':');
                            if (colonPos > 0)
                            {
                                messageLines.Add(trimmedLine.Substring(colonPos + 1).Trim());
                            }
                        }
                        else if (inMessage)
                        {
                            messageLines.Add(trimmedLine);
                        }
                    }

                    // 마지막 이벤트 처리
                    if (currentEvent.Count > 0 && messageLines.Count > 0)
                    {
                        currentEvent["Message"] = string.Join("\n", messageLines);
                    }

                    // 지속 시간 계산
                    CalculateEventDurations(events);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"이벤트 로그 조회 중 오류: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"상세 오류: {ex.InnerException.Message}");
                throw;
            }

            Console.WriteLine($"총 {events.Count}개의 이벤트를 찾았습니다.");
            return events.OrderBy(e => e.Timestamp).ToList();
        }

        private void CalculateEventDurations(List<ScreensaverEvent> events)
        {
            // 시간순으로 정렬
            var sortedEvents = events.OrderBy(e => e.Timestamp).ToList();

            // 4802(시작)-4803(종료) 쌍 찾기
            var openMap = new Dictionary<(string user, string pc), ScreensaverEvent>();
            foreach (var e in sortedEvents)
            {
                var key = (e.Username, e.ComputerName);
                if (e.EventId == 4802)
                    openMap[key] = e;      // 마지막 시작 이벤트로 갱신
                else if (e.EventId == 4803 && openMap.TryGetValue(key, out var startEvt))
                {
                    var duration = e.Timestamp - startEvt.Timestamp;
                    startEvt.Duration = e.Duration = duration;
                    openMap.Remove(key);
                }
            }
        }

        private static string ExtractUsernameFromMessage(string message)
        {
            try
            {
                var match = System.Text.RegularExpressions.Regex.Match(
                    message,
                    @"(?:계정 이름|Account Name):\s+(?:(?<domain>[^\\\s]+)\\)?(?<username>[^\s\r\n]+)"
                );

                if (match.Success)
                {
                    return match.Groups["username"].Value;
                }
            }
            catch { }

            return "Unknown";
        }

        public List<ScreensaverUsage> AnalyzeScreensaverUsage(IEnumerable<ScreensaverEvent> events)
        {
            var dict = new Dictionary<string, ScreensaverUsage>();

            // 모든 이벤트 출력 (디버깅용)
            Console.WriteLine("분석 중인 이벤트:");
            foreach (var e in events)
            {
                Console.WriteLine($"  ID: {e.EventId}, 시간: {e.Timestamp}, 사용자: {e.Username}, 컴퓨터: {e.ComputerName}");
            }

            foreach (var e in events.OrderBy(e => e.Timestamp))
            {
                if (!dict.TryGetValue(e.Username, out var usage))
                {
                    usage = new ScreensaverUsage(e.Username);
                    usage.ComputerName = e.ComputerName; // 컴퓨터 정보 설정
                    dict[e.Username] = usage;
                }

                switch (e.EventId)
                {
                    case 4802: // 화면보호기 시작
                        usage.ScreensaverPeriods.Add((e.Timestamp, null, e.EventId));
                        Console.WriteLine($"화면보호기 시작 추가: {e.Username}, {e.Timestamp}, {e.EventId}");
                        break;

                    case 4803: // 화면보호기 종료
                        // 매칭되는 시작 이벤트 찾기
                        var open = usage.ScreensaverPeriods
                            .Where(p => p.End == null && p.EventId == 4802) // 4802로 시작된 항목만 검색
                            .OrderBy(p => p.Start)
                            .FirstOrDefault();

                        if (open.Start != default)
                        {
                            int idx = usage.ScreensaverPeriods.IndexOf(open);
                            usage.ScreensaverPeriods[idx] = (open.Start, e.Timestamp, open.EventId);
                            Console.WriteLine($"화면보호기 종료 매칭: {e.Username}, {open.Start} ~ {e.Timestamp}");
                        }
                        else
                        {
                            // 4803 이벤트도 독립적으로 저장 (시작 이벤트 누락 시)
                            usage.ScreensaverPeriods.Add((e.Timestamp, e.Timestamp, e.EventId));
                            Console.WriteLine($"화면보호기 종료 이벤트 추가: {e.Username}, {e.Timestamp}, {e.EventId}");
                        }
                        break;
                }
            }

            foreach (var u in dict.Values)
            {
                // 화면보호기 활성화 횟수는 4802 이벤트 수로 계산
                u.ScreensaverActivationCount = u.ScreensaverPeriods.Count(p => p.EventId == 4802);

                // 총 지속 시간 계산 - null 안전하게 처리
                u.TotalScreensaverDuration = u.ScreensaverPeriods
                    .Where(p => p.End.HasValue) // 종료된 항목만
                    .Aggregate(TimeSpan.Zero, (acc, p) => acc + (p.End!.Value - p.Start)); // null이 아님을 명시(!.Value)
            }

            return dict.Values.ToList();
        }

        private static string GetEventUsername(EventRecord rec)
        {
            try
            {
                var xml = new XmlDocument(); xml.LoadXml(rec.ToXml());
                var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event");
                var node = xml.SelectSingleNode("//e:Data[@Name='TargetUserName']", ns) ?? xml.SelectSingleNode("//e:Data[@Name='SubjectUserName']", ns);
                if (node != null && !string.IsNullOrEmpty(node.InnerText)) return node.InnerText;
                return rec.Properties.Count > 1 ? rec.Properties[1].Value?.ToString() ?? "Unknown" : "Unknown";
            }
            catch { return "Unknown"; }
        }

        // 관리자 권한 확인 메서드
        public static bool IsAdministrator()
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
