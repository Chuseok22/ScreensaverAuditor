using System.Diagnostics;
using System.Security.Principal;
using System.Diagnostics.Eventing.Reader;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Xml;
using System.Drawing.Text;

class Program
{
    static int Main(string[] args)
    {
        try
        {
            // 초기 변수 설정

            // 실제 감사 정책 활성화, 이벤트 조회, 분석 기능 클래스
            var auditor = new ScreensaverAuditor();
            // --enable-policy 옵션이 없으면 기본값 false
            bool enablePolicy = false;
            // 기본 날짜 범위 설정
            // 시작 날짜: 7일 전, 종료 날짜: 오늘
            DateTime startDate = DateTime.Now.AddDays(-7);
            DateTime endDate = DateTime.Now;
            // 기본 CSV 파일명 --output 옵션이 없으면 기본 파일명 사용
            string outputPath = "ScreensaverEvents.csv";
            // 사용자 필터링을 위한 변수 --user 옵션이 없으면 null
            string? filterUsername = null;

            // 옵션 파싱
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    // --enable-policy 옵션
                    // 해당 옵션이 있으면 감사 정책 활성화 모드 작동
                    case "--enable-policy":
                        enablePolicy = true;
                        break;
                    // --start-date, --end-date을 "yyyy-MM-dd" 포맷으로 파싱
                    case "--start-date":
                        if (i + 1 < args.Length)
                        {
                            // 명시한 형식 ("YYYY-MM-DD")만 허용
                            if (DateTime.TryParseExact(args[i + 1], "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var sd))
                                startDate = sd;
                            else
                            {
                                Console.Error.WriteLine($"Error: --start-date '{args[i + 1]}' 형식 오류 (YYYY-MM-DD 형식으로 입력하세요.)");
                                return 3;
                            }
                        }
                        break;
                    case "--end-date":
                        if (i + 1 < args.Length)
                        {
                            // 명시한 형식 ("YYYY-MM-DD")만 허용
                            if (DateTime.TryParseExact(args[i + 1], "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var ed))
                                endDate = ed;
                            else
                            {
                                Console.Error.WriteLine($"Error: --end-date '{args[i + 1]}' 형식 오류 (YYYY-MM-DD 형식으로 입력하세요.)");
                                return 3;
                            }
                        }
                        break;
                    // CSV 파일 저장 옵션
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            outputPath = args[i + 1];
                            i++;
                        }
                        break;
                    // 사용자 필터링 옵션
                    case "--user":
                    case "--username":
                        if (i + 1 < args.Length)
                        {
                            filterUsername = args[++i];
                        }
                        else
                        {
                            Console.Error.WriteLine("Error: --user 옵션 뒤에 사용자 이름을 입력해야 합니다.");
                            return 4;
                        }
                        break;
                    // 도움말 옵션
                    case "--help":
                        ShowUsage();
                        return 0;
                    // 모르는 옵션의 경우 즉시 에러 출력 후 사용법 출력
                    default:
                        Console.WriteLine($"Error: '{args[i]}'는(은) 인식할 수 없는 옵션입니다.");
                        ShowUsage();
                        return 2;
                }
            }

            // --enable-policy 옵션이 활성화된 경우
            // 내부에서 관리자 권한 확인 후 IsAdministrator() 메서드로 확인
            if (enablePolicy)
                auditor.EnableAuditPolicy();

            var events = auditor.GetScreensaverEvents(startDate, endDate, filterUsername);
            var analysis = auditor.AnalyzeScreensaverUsage(events);

            // 결과 출력
            foreach (var usage in analysis)
            {
                Console.WriteLine($"사용자: {usage.Username}");
                Console.WriteLine($"화면보호기 실행 횟수: {usage.ScreensaverActivationCount}");
                Console.WriteLine($"총 화면보호기 지속 시간: {usage.TotalScreensaverDuration.TotalHours:F2} 시간");
                Console.WriteLine("----------------------------------------");
            }

            // CSV 파일로 저장
            {
                // 경로 유효성 검사 및 디렉토리 생성
                string directory = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                using (var writer = new StreamWriter(outputPath))
                {
                    writer.WriteLine("사용자,화면보호기 실행 횟수,총 지속 시간(시간)");
                    foreach (var usage in analysis)
                    {
                        // CSV 이스케이핑 처리
                        string escapedUsername = $"\"{usage.Username.Replace("\"", "\"\"")}\"";
                        writer.WriteLine($"{escapedUsername},{usage.ScreensaverActivationCount},{usage.TotalScreensaverDuration.TotalHours:F2}");
                    }
                }
                Console.WriteLine($"결과가 {outputPath}에 저장되었습니다.");
            }

            return 0;
        }
        catch (UnauthorizedAccessException)
        {
            Console.Error.WriteLine("Error: 관리자 권한이 필요합니다. EXE를 우클릭 -> '관리자 권한으로 실행' 해주세요.");
            return 1;
        }
        catch (Exception ex)
        {
            // 일반 오류는 표준 에러 스트림으로
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 2;
        }
    }

    static void ShowUsage()
    {
        Console.WriteLine("사용법: ScreensaverAuditor.exe [옵션]");
        Console.WriteLine("옵션:");
        Console.WriteLine("  --enable-policy       감사 정책 활성화");
        Console.WriteLine("  --start-date <날짜>  시작 날짜 (형식: YYYY-MM-DD) 기본값: 7일전");
        Console.WriteLine("  --end-date <날짜>    종료 날짜 (형식: YYYY-MM-DD) 기본값: 오늘");
        Console.WriteLine("  --output <파일경로>  결과를 CSV 파일로 저장");
        Console.WriteLine("  --user <사용자명>     특정 사용자의 이벤트만 필터링");
        Console.WriteLine("  --help               도움말 표시");
    }
}

public class ScreensaverEvent
{
    public DateTime Timestamp { get; set; }
    public string EventId { get; }
    public string Username { get; }
    public string ComputerName { get; }

    public ScreensaverEvent(DateTime timestamp, string eventId, string username, string computerName)
    {
        Timestamp = timestamp;
        EventId = eventId;
        Username = username;
        ComputerName = computerName;
    }
}

public class ScreensaverUsage
{
    public string Username { get; }
    public int ScreensaverActivationCount { get; set; }
    public TimeSpan TotalScreensaverDuration { get; set; }
    public List<(DateTime Start, DateTime? End)> ScreensaverPeriods { get; }

    public ScreensaverUsage(string username)
    {
        Username = username;
        ScreensaverPeriods = new List<(DateTime Start, DateTime? End)>();
    }
}

public class ScreensaverAuditor
{
    // 감사할 하위 범주 상수 선언
    private const string AuditSubcategory = "기타 로그온/로그오프 이벤트";
    public void EnableAuditPolicy()
    {
        try
        {
            if (!IsAdministrator())
            {
                throw new UnauthorizedAccessException("이 작업을 수행하려면 관리자 권한이 필요합니다.");
            }

            // 1) 감사 정책 활성화
            var setInfo = new ProcessStartInfo
            {
                FileName = "auditpol.exe",
                Arguments = $"/set /subcategory:\"{AuditSubcategory}\" /success:enable /failure:enable",
                UseShellExecute = false
            };
            using (var process = Process.Start(setInfo))
            {
                if (process == null)
                {
                    throw new Exception("프로세스를 시작할 수 없습니다.");
                }

                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    throw new Exception("감사 정책을 설정하는 데 실패했습니다.");
                }
            }

            // 2) 설정 확인 명령 실행
            var getInfo = new ProcessStartInfo
            {
                FileName = "auditpol.exe",
                Arguments = $"/get /subcategory:\"{AuditSubcategory}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true
            };
            using (var checkProcess = Process.Start(getInfo))
            {
                if (checkProcess == null)
                {
                    throw new Exception("프로세스를 시작할 수 없습니다.");
                }

                string output = checkProcess.StandardOutput.ReadToEnd();
                checkProcess.WaitForExit();
                Console.WriteLine("감사 정책 확인 결과:");
                Console.WriteLine(output);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"감사 정책 설정 중 오류 발생: {ex.Message}");
            throw;
        }
    }

    public List<ScreensaverEvent> GetScreensaverEvents(DateTime startDate, DateTime endDate, string? username = null)
    {
        var events = new List<ScreensaverEvent>();
        var query = $@"*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] 
            and (EventID=4800 or EventID=4801 or EventID=4802 or EventID=4803)
            and TimeCreated[@SystemTime>='{startDate:o}' and @SystemTime<='{endDate:o}']]]";

        using (var eventLog = new EventLogReader(
            new EventLogQuery("Security", PathType.LogName, query)))
        {
            EventRecord record;
            while ((record = eventLog.ReadEvent()) != null)
            {
                using (record)
                {
                    var evt = new ScreensaverEvent(
                        record.TimeCreated ?? DateTime.MinValue,
                        record.Id.ToString(),
                        GetEventUsername(record),
                        record.MachineName
                    );

                    if (username == null || evt.Username.Contains(username, StringComparison.OrdinalIgnoreCase))
                    {
                        events.Add(evt);
                    }
                }
            }
        }

        return events;
    }

    public List<ScreensaverUsage> AnalyzeScreensaverUsage(List<ScreensaverEvent> events)
    {
        var usageByUser = new Dictionary<string, ScreensaverUsage>();

        foreach (var evt in events.OrderBy(e => e.Timestamp))
        {
            if (!usageByUser.ContainsKey(evt.Username))
            {
                usageByUser[evt.Username] = new ScreensaverUsage(evt.Username);
            }

            var usage = usageByUser[evt.Username];

            switch (evt.EventId)
            {
                case "4802": // 화면보호기 시작
                    usage.ScreensaverPeriods.Add((evt.Timestamp, null));
                    break;
                case "4803": // 화면보호기 종료
                    // 열린 세션들을 시작 오름차순으로 정렬한 뒤, 가장 오래된 세션만 종료
                    var openPeriods = usage.ScreensaverPeriods
                        .Where(p => p.End == null)
                        .OrderBy(p => p.Start)
                        .ToList();
                    if (openPeriods.Any())
                    {
                        var oldest = openPeriods.First();
                        int index = usage.ScreensaverPeriods.IndexOf(oldest);
                        usage.ScreensaverPeriods[index] = (oldest.Start, evt.Timestamp);
                    }
                    break;
            }
        }

        return usageByUser.Select(kvp =>
        {
            var usage = kvp.Value;
            usage.ScreensaverActivationCount = usage.ScreensaverPeriods.Count;
            usage.TotalScreensaverDuration = CalculateTotalDuration(usage.ScreensaverPeriods);
            return usage;
        }).ToList();
    }

    private TimeSpan CalculateTotalDuration(List<(DateTime Start, DateTime? End)> periods)
    {
        TimeSpan total = TimeSpan.Zero;
        foreach (var period in periods)
        {
            if (period.End.HasValue)
            {
                total += period.End.Value - period.Start;
            }
            else
            {
                total += DateTime.Now - period.Start;
            }
        }
        return total;
    }

    private string GetEventUsername(EventRecord record)
    {
        try
        {
            // XML 로드 및 네임스페이스 설정
            var xml = new XmlDocument();
            xml.LoadXml(record.ToXml());
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event");

            // 우선순위 1: TargetUserName
            var node = xml.SelectSingleNode("//e:Data[@Name='TargetUserName']", ns)
                    ?? xml.SelectSingleNode("//e:Data[@Name='SubjectUserName']", ns);

            if (node != null && !string.IsNullOrEmpty(node.InnerText))
                return node.InnerText;

            // 최후의 보루: 기존 속성 기반
            if (record.Properties.Count > 1)
                return record.Properties[1].Value?.ToString() ?? "Unknown";

            return "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    private bool IsAdministrator()
    {
        // WindowsIdentity는 IDisposable이므로 using 블록으로 감싸 리소스 누수 방지
        using (var identity = WindowsIdentity.GetCurrent())
        {
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
