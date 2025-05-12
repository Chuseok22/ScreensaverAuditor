// Utils/CommandLineParser.cs
using System;
using ScreensaverAuditor.Models;

namespace ScreensaverAuditor.Utils
{
    public static class CommandLineParser
    {
        // 기본적으로 7일 전부터 오늘까지의 날짜를 조회합니다.
        private const int DAYS_TO_LOOK_BACK = 7;

        public static CommandLineOptions ParseCommandLineOptions(string[] args)
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

        public static void ShowUsage()
        {
            Console.WriteLine("사용법: ScreensaverAuditor.exe [옵션]");
            Console.WriteLine("  --enable-policy          감사 정책 활성화");
            Console.WriteLine("  --start-date YYYY-MM-DD  조회 시작 날짜 (기본: 7일 전)");
            Console.WriteLine("  --end-date   YYYY-MM-DD  조회 종료 날짜 (기본: 오늘)");
            Console.WriteLine("  --user <이름>            특정 사용자만 필터링");
            Console.WriteLine("  --output <파일경로>      결과 저장 경로 (기본: ScreensaverEvents.xlsx)");
            Console.WriteLine("  --help                   도움말 표시");
        }
    }
}
