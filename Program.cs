using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Security.Principal;
using System.Xml;
using ClosedXML.Excel;

namespace ScreensaverAuditor
{
    public class Program
    {
        private const string DEFAULT_EXCEL_FILE = "ScreensaverEvents.xlsx";
        private const int DAYS_TO_LOOK_BACK = 7;

        private static readonly ExcelExporter excelExporter = new();
        private static readonly ScreensaverAuditor auditor = new();

        static int Main(string[] args)
        {
            try
            {
                var options = ParseCommandLineOptions(args);

                if (options.ShowHelp)
                {
                    ShowUsage();
                    return 0;
                }

                if (options.EnablePolicy)
                {
                    auditor.EnableAuditPolicy();
                }

                var events = auditor.GetScreensaverEvents(options.StartDate, options.EndDate, options.Username);
                var analysis = auditor.AnalyzeScreensaverUsage(events);

                DisplayResults(analysis);
                SaveEventDetails(events, options.OutputPath ?? DEFAULT_EXCEL_FILE); // 이벤트 상세 정보 저장

                if (args.Length == 0) WaitForKey();
                return 0;
            }
            catch (Exception ex)
            {
                HandleException(ex, args.Length == 0);
                return GetErrorCode(ex);
            }
        }

        private static void SaveEventDetails(List<ScreensaverEvent> events, string outputPath)
        {
            excelExporter.SaveToExcel(outputPath, events);
            Console.WriteLine($"결과가 '{outputPath}'에 저장되었습니다.");
        }

        private static CommandLineOptions ParseCommandLineOptions(string[] args)
        {
            var options = new CommandLineOptions
            {
                StartDate = DateTime.Now.AddDays(-DAYS_TO_LOOK_BACK),
                EndDate = DateTime.Now
            };

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "--enable-policy":
                        options.EnablePolicy = true;
                        break;

                    case "--start-date":
                        if (!TryParseDate(args, ref i, out var startDate))
                            throw new ArgumentException("시작 날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)");
                        options.StartDate = startDate;
                        break;

                    case "--end-date":
                        if (!TryParseDate(args, ref i, out var endDate))
                            throw new ArgumentException("종료 날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)");
                        options.EndDate = endDate;
                        break;

                    case "--user":
                    case "--username":
                        if (i + 1 >= args.Length)
                            throw new ArgumentException("사용자 이름이 지정되지 않았습니다");
                        options.Username = args[++i];
                        break;

                    case "--output":
                        if (i + 1 >= args.Length)
                            throw new ArgumentException("출력 파일 경로가 지정되지 않았습니다");
                        options.OutputPath = args[++i];
                        break;

                    case "--help":
                        options.ShowHelp = true;
                        return options;

                    default:
                        throw new ArgumentException($"인식할 수 없는 옵션: {args[i]}");
                }
            }

            return options;
        }

        private static bool TryParseDate(string[] args, ref int index, out DateTime result)
        {
            result = DateTime.Now;
            if (index + 1 >= args.Length) return false;

            return DateTime.TryParseExact(
                args[++index],
                "yyyy-MM-dd",
                null,
                System.Globalization.DateTimeStyles.None,
                out result);
        }

        private static void DisplayResults(List<ScreensaverUsage> analysis)
        {
            foreach (var usage in analysis)
            {
                Console.WriteLine($"사용자: {usage.Username}");
                Console.WriteLine($"화면보호기 실행 횟수: {usage.ScreensaverActivationCount}");
                Console.WriteLine($"총 지속 시간: {usage.TotalScreensaverDuration.TotalHours:F2} 시간");
                Console.WriteLine("───────────────────────────────────────");
            }
        }

        private static void ShowUsage()
        {
            Console.WriteLine("사용법: ScreensaverAuditor.exe [옵션]");
            Console.WriteLine("  --enable-policy          감사 정책 활성화");
            Console.WriteLine("  --start-date YYYY-MM-DD  조회 시작 날짜 (기본: 7일 전)");
            Console.WriteLine("  --end-date   YYYY-MM-DD  조회 종료 날짜 (기본: 오늘)");
            Console.WriteLine("  --user <이름>            특정 사용자만 필터링");
            Console.WriteLine("  --output <파일경로>      결과 저장 경로 (기본: ScreensaverEvents.xlsx)");
            Console.WriteLine("  --help                   도움말 표시");
        }

        private static void WaitForKey()
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void HandleException(Exception ex, bool waitForKey)
        {
            string message = ex switch
            {
                UnauthorizedAccessException _ => "Error: 관리자 권한이 필요합니다. EXE를 우클릭 → '관리자 권한으로 실행'",
                ArgumentException _ => $"Error: {ex.Message}",
                _ => $"Error: 예기치 않은 오류가 발생했습니다. {ex.Message}"
            };

            Console.Error.WriteLine(message);
            if (waitForKey) WaitForKey();
        }

        private static int GetErrorCode(Exception ex) => ex switch
        {
            UnauthorizedAccessException => 1,
            ArgumentException => 2,
            _ => 3
        };
    }

    public class CommandLineOptions
    {
        public bool EnablePolicy { get; set; }
        public bool ShowHelp { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string? Username { get; set; }
        public string? OutputPath { get; set; }
    }

    // ───────────────────────── XLSX 저장 클래스 ─────────────────────────
    public class ExcelExporter
    {
        private void EnsureValidPath(string outputPath)
        {
            string finalPath = !outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? outputPath + ".xlsx"
                : outputPath;

            if (File.Exists(finalPath))
            {
                try { File.Delete(finalPath); }
                catch (IOException) { throw new IOException("Excel 파일이 열려 있습니다. 닫고 다시 시도하세요."); }
            }

            var dir = Path.GetDirectoryName(finalPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private void AddHeaders(IXLWorksheet worksheet)
        {
            string[] headers = {
        "시간",
        "이벤트 ID",
        "PC 관리번호",
        "메시지",
        "계정 이름",
        "계정 도메인",
        "보안 ID",
        "로그온 ID",
        "세션 ID",
        "지속시간(분)"
    };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
            }
        }

        private void AddData(IXLWorksheet worksheet, List<ScreensaverEvent> events)
        {
            int row = 2;
            foreach (var evt in events)
            {
                AddEventRow(worksheet, row++, evt);
            }
        }

        private void AddEventRow(IXLWorksheet worksheet, int row, ScreensaverEvent evt)
        {
            // 기본 정보
            var dateCell = worksheet.Cell(row, 1);
            dateCell.Value = evt.Timestamp;
            dateCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";

            worksheet.Cell(row, 2).Value = evt.EventId;
            worksheet.Cell(row, 3).Value = evt.ComputerName;
            worksheet.Cell(row, 4).Value = evt.Message;
            worksheet.Cell(row, 5).Value = evt.Username;
            worksheet.Cell(row, 6).Value = evt.AccountDomain;
            worksheet.Cell(row, 7).Value = evt.SecurityId;
            worksheet.Cell(row, 8).Value = evt.LogonId;
            worksheet.Cell(row, 9).Value = evt.SessionId;

            // 지속시간 (분)
            var durationCell = worksheet.Cell(row, 10);
            if (evt.Duration.HasValue)
            {
                durationCell.Value = Math.Round(evt.Duration.Value.TotalMinutes, 2);
            }
            else
            {
                durationCell.Value = string.Empty;
            }
        }

        private void ApplyFormatting(IXLWorksheet worksheet)
        {
            var headerRange = worksheet.Range(1, 1, 1, 10);  // 10개 열
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            worksheet.SheetView.FreezeRows(1);
            worksheet.Columns().AdjustToContents();
        }

        // ExcelExporter 클래스의 SaveToExcel 메서드 수정
        public void SaveToExcel(string outputPath, List<ScreensaverEvent> events)
        {
            EnsureValidPath(outputPath);

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Events");

            AddHeaders(worksheet);
            AddData(worksheet, events);
            ApplyFormatting(worksheet);

            workbook.SaveAs(outputPath);
        }
    }

    // ───────────────────────── 데이터 모델 ─────────────────────────
    public class ScreensaverEvent
    {
        public DateTime Timestamp { get; }
        public int EventId { get; }
        public string ComputerName { get; }
        public string Message { get; }
        public string Username { get; }
        public string AccountDomain { get; }
        public string SecurityId { get; }
        public string LogonId { get; }
        public string SessionId { get; }
        public TimeSpan? Duration { get; set; } // 지속시간 (계산됨)

        public ScreensaverEvent(
            DateTime ts,
            int id,
            string computer,
            string message,
            string username,
            string domain,
            string securityId,
            string logonId,
            string sessionId)
        {
            Timestamp = ts;
            EventId = id;
            ComputerName = computer;
            Message = message;
            Username = username;
            AccountDomain = domain;
            SecurityId = securityId;
            LogonId = logonId;
            SessionId = sessionId;
            Duration = null;
        }
    }

    public class ScreensaverUsage
    {
        public string Username { get; }
        public string ComputerName { get; set; } = ""; // 컴퓨터 정보 추가
        public int ScreensaverActivationCount { get; set; }
        public TimeSpan TotalScreensaverDuration { get; set; }
        public List<(DateTime Start, DateTime? End, int EventId)> ScreensaverPeriods { get; } // EventId를 정수형으로 변경

        public ScreensaverUsage(string user)
        {
            Username = user;
            ScreensaverPeriods = new List<(DateTime, DateTime?, int)>();
        }
    }

    // ───────────────────────── 핵심 로직 ─────────────────────────
    public class ScreensaverAuditor
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
                    var safeUser = user.Replace("`", "``").Replace("'", "''");
                    command += $" | Where-Object {{ $_.Message -like '*{safeUser}*' }}";
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
            for (int i = 0; i < sortedEvents.Count - 1; i++)
            {
                if (sortedEvents[i].EventId == 4802 && sortedEvents[i + 1].EventId == 4803)
                {
                    // 같은 사용자 && 같은 PC의 이벤트
                    if (sortedEvents[i].Username == sortedEvents[i + 1].Username &&
                        sortedEvents[i].ComputerName == sortedEvents[i + 1].ComputerName)
                    {
                        TimeSpan duration = sortedEvents[i + 1].Timestamp - sortedEvents[i].Timestamp;
                        sortedEvents[i].Duration = duration; // 시작 이벤트에 지속시간 설정
                        sortedEvents[i + 1].Duration = duration; // 종료 이벤트에도 지속시간 설정
                    }
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
                // 화면보호기 활성화 횟수는 4802와 4803 이벤트의 합계로 계산
                u.ScreensaverActivationCount = u.ScreensaverPeriods.Count(p => p.EventId == 4802 || p.EventId == 4803);

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

        private static bool IsAdministrator()
        {
            using var id = WindowsIdentity.GetCurrent();
            return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
