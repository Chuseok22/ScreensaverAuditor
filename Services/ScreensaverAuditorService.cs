// Services/ScreensaverAuditorService.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Xml;
using System.Diagnostics.Eventing.Reader;
using ScreensaverAuditor.Models;
using System.Text.Json;

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

                command += " | ConvertTo-Json -Depth 4"; // Depth 매개변수 추가하여 복잡한 객체도 직렬화되도록 함

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

                Console.WriteLine("PowerShell 명령 실행 중..."); 

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

                    // 출력 파싱 부분을 JSON 기반으로 변경
                    Console.WriteLine("이벤트 로그 파싱 시작...");

                    // JSON 출력이 비어있는지 확인
                    if (string.IsNullOrWhiteSpace(output) || output.Trim() == "null" || output.Trim() == "[]")
                    {
                        Console.WriteLine("지정된 기간에 이벤트가 없습니다.");
                        return events;
                    }

                    try
                    {
                        // 단일 객체인지 배열인지 확인하여 적절히 처리
                        JsonDocument doc = JsonDocument.Parse(output);
                        if (doc.RootElement.ValueKind == JsonValueKind.Array)
                        {
                            // 배열 처리
                            foreach (JsonElement jsonEvent in doc.RootElement.EnumerateArray())
                            {
                                ProcessJsonEvent(jsonEvent, events);
                            }
                        }
                        else if (doc.RootElement.ValueKind == JsonValueKind.Object)
                        {
                            // 단일 객체 처리
                            ProcessJsonEvent(doc.RootElement, events);
                        }
                        else
                        {
                            Console.WriteLine($"예상치 못한 JSON 형식: {doc.RootElement.ValueKind}");
                            Console.WriteLine($"JSON 내용: {output.Substring(0, Math.Min(output.Length, 200))}..."); // 일부만 출력
                        }
                    }
                    catch (JsonException ex)
                    {
                        Console.WriteLine($"JSON 파싱 오류: {ex.Message}");
                        Console.WriteLine($"JSON 내용: {output.Substring(0, Math.Min(output.Length, 200))}...");
                    }

                    // 지속 시간 계산
                    if (events.Count > 0)
                    {
                        CalculateEventDurations(events);
                    }
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

        // JSON 이벤트 처리 메서드
        private void ProcessJsonEvent(JsonElement jsonEvent, List<ScreensaverEvent> events)
        {
            try
            {
                // 필수 속성이 있는지 확인
                if (!jsonEvent.TryGetProperty("TimeCreated", out var timeCreatedElement) ||
                    !jsonEvent.TryGetProperty("Id", out var idElement) ||
                    !jsonEvent.TryGetProperty("MachineName", out var machineElement) ||
                    !jsonEvent.TryGetProperty("Message", out var messageElement))
                {
                    Console.WriteLine("필수 속성이 누락된 이벤트를 건너뜁니다.");
                    return;
                }

                // 추가 속성 확인
                string taskDisplayName = string.Empty;
                string keywords = string.Empty;
                string activityId = string.Empty;
                string providerName = string.Empty;
                
                if (jsonEvent.TryGetProperty("TaskDisplayName", out var taskElement))
                    taskDisplayName = taskElement.ToString();
                
                if (jsonEvent.TryGetProperty("KeywordsDisplayNames", out var keywordsElement) && 
                    keywordsElement.ValueKind == JsonValueKind.Array)
                {
                    keywords = string.Join(", ", keywordsElement.EnumerateArray().Select(e => e.ToString()));
                }
                
                if (jsonEvent.TryGetProperty("ActivityId", out var activityElement))
                    activityId = activityElement.ToString();
                    
                if (jsonEvent.TryGetProperty("ProviderName", out var providerElement))
                    providerName = providerElement.ToString();

                // 시간 파싱 - 특별한 형식 확인 (/Date(timestamp)/)
                DateTime timeCreated;
                string timeCreatedStr = timeCreatedElement.ToString();
                
                if (timeCreatedStr.StartsWith("/Date(") && timeCreatedStr.EndsWith(")/"))
                {
                    // Microsoft JSON 날짜 형식 처리
                    string timestampStr = timeCreatedStr.Substring(6, timeCreatedStr.Length - 8);
                    if (long.TryParse(timestampStr.Split('+', '-')[0], out long timestamp))
                    {
                        // Unix 에폭 시간 (밀리초)를 DateTime으로 변환
                        timeCreated = DateTimeOffset.FromUnixTimeMilliseconds(timestamp).DateTime;
                    }
                    else
                    {
                        Console.WriteLine($"Microsoft JSON 날짜 형식 파싱 오류: {timeCreatedStr}");
                        return;
                    }
                }
                else if (!DateTime.TryParse(timeCreatedStr, out timeCreated))
                {
                    Console.WriteLine($"시간 파싱 오류: {timeCreatedStr}");
                    return;
                }

                // ID 파싱
                if (!int.TryParse(idElement.ToString(), out var id))
                {
                    Console.WriteLine($"ID 파싱 오류: {idElement}");
                    return;
                }

                string computerName = machineElement.ToString();
                string message = messageElement.ToString();

                // 사용자 정보 추출
                var userInfo = ExtractUserInfoFromMessage(message);

                // 이벤트 종류 결정 (4802: 활성화, 4803: 비활성화)
                string eventType = id == 4802 ? "화면보호기 시작" : "화면보호기 종료";

                // 이벤트 생성 및 추가
                var evt = new ScreensaverEvent(
                    timeCreated,
                    id,
                    computerName,
                    message,
                    userInfo.Username,
                    userInfo.Domain,
                    userInfo.SecurityId,
                    userInfo.LogonId,
                    userInfo.SessionId,
                    taskDisplayName,
                    activityId,
                    providerName,
                    keywords);

                events.Add(evt);
                Console.WriteLine($"이벤트 추가: ID={id}, 시간={timeCreated:yyyy-MM-dd HH:mm:ss}, 사용자={userInfo.Username}, PC={computerName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"이벤트 처리 중 오류 발생: {ex.Message}");
            }
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

        // 메시지에서 사용자 정보를 추출하는 메서드
        private UserInfo ExtractUserInfoFromMessage(string message)
        {
            var userInfo = new UserInfo
            {
                Username = "Unknown",
                Domain = "Unknown",
                SecurityId = "Unknown",
                LogonId = "Unknown",
                SessionId = "Unknown"
            };

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
                            userInfo.Domain = accountParts[0].Trim();
                            userInfo.Username = accountParts[1].Trim();
                        }
                        else
                        {
                            userInfo.Username = parts[1].Trim();
                        }
                    }
                }
                else if (msgTrimmed.StartsWith("계정 도메인:") || msgTrimmed.StartsWith("Account Domain:"))
                {
                    var parts = msgTrimmed.Split(':', 2);
                    if (parts.Length > 1)
                    {
                        userInfo.Domain = parts[1].Trim();
                    }
                }
                else if (msgTrimmed.StartsWith("보안 ID:") || msgTrimmed.StartsWith("Security ID:"))
                {
                    var parts = msgTrimmed.Split(':', 2);
                    if (parts.Length > 1)
                    {
                        userInfo.SecurityId = parts[1].Trim();
                    }
                }
                else if (msgTrimmed.StartsWith("로그온 ID:") || msgTrimmed.StartsWith("Logon ID:"))
                {
                    var parts = msgTrimmed.Split(':', 2);
                    if (parts.Length > 1)
                    {
                        userInfo.LogonId = parts[1].Trim();
                    }
                }
                else if (msgTrimmed.StartsWith("세션 ID:") || msgTrimmed.StartsWith("Session ID:"))
                {
                    var parts = msgTrimmed.Split(':', 2);
                    if (parts.Length > 1)
                    {
                        userInfo.SessionId = parts[1].Trim();
                    }
                }
            }

            return userInfo;
        }

        // 사용자 정보를 담는 클래스
        private class UserInfo
        {
            public string Username { get; set; } = "Unknown";
            public string Domain { get; set; } = "Unknown";
            public string SecurityId { get; set; } = "Unknown";
            public string LogonId { get; set; } = "Unknown";
            public string SessionId { get; set; } = "Unknown";
        }

        public List<ScreensaverUsage> AnalyzeScreensaverUsage(IEnumerable<ScreensaverEvent> events)
        {
            var dict = new Dictionary<string, ScreensaverUsage>();

            // 모든 이벤트 출력 (디버깅용)
            Console.WriteLine("분석 중인 이벤트:");
            foreach (var e in events)
            {
                Console.WriteLine($"  ID: {e.EventId}, 시간: {e.Timestamp}, 사용자: {e.Username}, 컴퓨터: {e.ComputerName}"); // TODO: 추후 프로덕션에서 제거
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
                        Console.WriteLine($"화면보호기 시작 추가: {e.Username}, {e.Timestamp}, {e.EventId}"); // TODO: 추후 프로덕션에서 제거
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

        // 관리자 권한 확인 메서드
        public static bool IsAdministrator()
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
