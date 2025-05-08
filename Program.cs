using System.Diagnostics;
using System.Security.Principal;
using System.Diagnostics.Eventing.Reader;

class Program
{
    static int Main(string[] args)
    {
        try {
            var auditor = new ScreensaverAuditor();

            bool enablePolicy = false;
            DateTime startDate = DateTime.Now.AddDays(-7);
            DateTime endDate = DateTime.Now;
            string? outputPath = null;

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--enable-policy":
                        enablePolicy = true;
                        break;
                    case "--start-date":
                        if (i + 1 < args.Length && DateTime.TryParse(args[i + 1], out var sd))
                        {
                            startDate = sd;
                            i++;
                        }
                        break;
                    case "--end-date":
                        if (i + 1 < args.Length && DateTime.TryParse(args[i + 1], out var ed))
                        {
                            endDate = ed;
                            i++;
                        }
                        break;
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            outputPath = args[i + 1];
                            i++;
                        }
                        break;
                    case "--help":
                        ShowUsage();
                        return 0;
                }
            }

            if (enablePolicy)
                auditor.EnableAuditPolicy();

            var events = auditor.GetScreensaverEvents(startDate, endDate);
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
            if (!string.IsNullOrEmpty(outputPath))
            {
                using (var writer = new StreamWriter(outputPath))
                {
                    writer.WriteLine("사용자,화면보호기 실행 횟수,총 지속 시간(시간)");
                    foreach (var usage in analysis)
                    {
                        writer.WriteLine($"{usage.Username},{usage.ScreensaverActivationCount},{usage.TotalScreensaverDuration.TotalHours:F2}");
                    }
                }
                Console.WriteLine($"결과가 {outputPath}에 저장되었습니다.");
            }

            return 0;
        } catch (UnauthorizedAccessException)
        {
            Console.Error.WriteLine("Error: 관리자 권한이 필요합니다. EXE를 우클릭 -> '관리자 권한으로 실행' 해주세요.");
            return 1;
        }
        catch (Exception ex)
        {
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
    public void EnableAuditPolicy()
    {
        try
        {
            if (!IsAdministrator())
            {
                throw new UnauthorizedAccessException("이 작업을 수행하려면 관리자 권한이 필요합니다.");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "auditpol.exe";
            startInfo.Arguments = "/set /subcategory:\"Other System Events\" /success:enable /failure:enable";
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.Verb = "runas";

            Process? process = Process.Start(startInfo);
            if (process == null)
            {
                throw new Exception("프로세스를 시작할 수 없습니다.");
            }

            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                throw new Exception("감사 정책 활성화 실패");
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
                    var lastPeriod = usage.ScreensaverPeriods.FindLast(p => p.End == null);
                    if (lastPeriod.Start != default)
                    {
                        int index = usage.ScreensaverPeriods.IndexOf(lastPeriod);
                        usage.ScreensaverPeriods[index] = (lastPeriod.Start, evt.Timestamp);
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
            if (record.Properties.Count > 1)
            {
                return record.Properties[1].Value?.ToString() ?? "Unknown";
            }
            return "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    private bool IsAdministrator()
    {
        using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
        {
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
