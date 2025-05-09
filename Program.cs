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
                SaveResults(analysis, options.OutputPath ?? DEFAULT_EXCEL_FILE);

                if (args.Length == 0) WaitForKey();
                return 0;
            }
            catch (Exception ex)
            {
                HandleException(ex, args.Length == 0);
                return GetErrorCode(ex);
            }
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

        private static void SaveResults(List<ScreensaverUsage> analysis, string outputPath)
        {
            excelExporter.SaveToExcel(outputPath, analysis);
            Console.WriteLine($"결과가 '{outputPath}'에 저장되었습니다.");
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
        public void SaveToExcel(string outputPath, List<ScreensaverUsage> analysis)
        {
            EnsureValidPath(outputPath);

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Usage");

            AddHeaders(worksheet);
            AddData(worksheet, analysis);
            ApplyFormatting(worksheet);

            workbook.SaveAs(outputPath);
        }

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
            string[] headers = { "사용자 정보", "날짜 및 시간", "이벤트 ID", "화면보호기 실행횟수", "총 지속 시간 (시간)" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
            }
        }

        private void AddData(IXLWorksheet worksheet, List<ScreensaverUsage> analysis)
        {
            int row = 2;
            foreach (var usage in analysis)
            {
                foreach (var period in usage.ScreensaverPeriods)
                {
                    AddRow(worksheet, row++, usage, period);
                }
            }
        }

        private void AddRow(IXLWorksheet worksheet, int row, ScreensaverUsage usage, (DateTime Start, DateTime? End) period)
        {
            string eventId = period.End.HasValue ? "4803" : "4802";
            double duration = period.End.HasValue
                ? (period.End.Value - period.Start).TotalHours
                : (DateTime.Now - period.Start).TotalHours;

            worksheet.Cell(row, 1).Value = usage.Username;

            var dateCell = worksheet.Cell(row, 2);
            dateCell.Value = period.Start;
            dateCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";

            worksheet.Cell(row, 3).Value = eventId;
            worksheet.Cell(row, 4).Value = usage.ScreensaverActivationCount;

            var durationCell = worksheet.Cell(row, 5);
            durationCell.Value = Math.Round(duration, 2);
            durationCell.Style.NumberFormat.Format = "#,##0.00";
        }

        private void ApplyFormatting(IXLWorksheet worksheet)
        {
            var headerRange = worksheet.Range(1, 1, 1, 5);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            worksheet.SheetView.FreezeRows(1);
            worksheet.Columns().AdjustToContents();
        }
    }

    // ───────────────────────── 데이터 모델 ─────────────────────────
    public class ScreensaverEvent
    {
        public DateTime Timestamp { get; }
        public string EventId { get; }
        public string Username { get; }
        public string ComputerName { get; }

        public ScreensaverEvent(DateTime ts, string id, string user, string comp)
        { Timestamp = ts; EventId = id; Username = user; ComputerName = comp; }
    }

    public class ScreensaverUsage
    {
        public string Username { get; }
        public int ScreensaverActivationCount { get; set; }
        public TimeSpan TotalScreensaverDuration { get; set; }
        public List<(DateTime Start, DateTime? End)> ScreensaverPeriods { get; }

        public ScreensaverUsage(string user)
        { Username = user; ScreensaverPeriods = new(); }
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
                // 임시 CSV 파일 경로
                string tempCsv = Path.Combine(Path.GetTempPath(), $"screensaver_events_{DateTime.Now:yyyyMMddHHmmss}.csv");

                // PowerShell 명령어 구성
                string command = $@"Get-WinEvent -FilterHashtable @{{
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

                command += " | Select-Object TimeCreated,Id,Message | Export-Csv -Path '" + tempCsv.Replace("'", "''") + "' -Encoding UTF8 -NoTypeInformation";

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

                using (var process = Process.Start(startInfo))
                {
                    if (process == null) throw new Exception("PowerShell 실행 실패");

                    // 오류 출력 확인
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        throw new Exception($"PowerShell 오류: {error}");
                    }

                    // CSV 파일 읽기
                    if (File.Exists(tempCsv))
                    {
                        using (var reader = new StreamReader(tempCsv))
                        {
                            // 헤더 건너뛰기
                            reader.ReadLine();

                            using var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader);
                            parser.HasFieldsEnclosedInQuotes = true;
                            parser.SetDelimiters(",");
                            while (!parser.EndOfData)
                            {
                                var parts = parser.ReadFields();
                                if (parts?.Length >= 3)
                                {
                                    // CSV 데이터 파싱
                                    var timestamp = DateTime.Parse(parts[0].Trim('"'));
                                    var eventId = parts[1].Trim('"');
                                    var message = parts[2].Trim('"');

                                    // 사용자 이름 추출 (Message에서)
                                    string username = ExtractUsernameFromMessage(message);

                                    var evt = new ScreensaverEvent(
                                        timestamp,
                                        eventId,
                                        username,
                                        Environment.MachineName);

                                    events.Add(evt);
                                    Console.WriteLine($"이벤트 발견: ID={evt.EventId}, 시간={evt.Timestamp:yyyy-MM-dd HH:mm:ss}, 사용자={evt.Username}");
                                }
                            }
                        }

                        // 임시 파일 삭제
                        try { File.Delete(tempCsv); }
                        catch { /* 무시 */ }
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
            foreach (var e in events.OrderBy(e => e.Timestamp))
            {
                if (!dict.TryGetValue(e.Username, out var usage))
                {
                    usage = new ScreensaverUsage(e.Username);
                    dict[e.Username] = usage;
                }
                switch (e.EventId)
                {
                    case "4802":
                        usage.ScreensaverPeriods.Add((e.Timestamp, null));
                        break;
                    case "4803":
                        var open = usage.ScreensaverPeriods.FirstOrDefault(p => p.End == null);
                        if (open.Start != default)
                        {
                            int idx = usage.ScreensaverPeriods.IndexOf(open);
                            usage.ScreensaverPeriods[idx] = (open.Start, e.Timestamp);
                        }
                        break;
                }
            }
            foreach (var u in dict.Values)
            {
                u.ScreensaverActivationCount = u.ScreensaverPeriods.Count;
                u.TotalScreensaverDuration = u.ScreensaverPeriods.Aggregate(TimeSpan.Zero, (acc, p) => acc + ((p.End ?? DateTime.Now) - p.Start));
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
