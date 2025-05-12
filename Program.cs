using System;
using System.Collections.Generic;
using System.Diagnostics;
using ScreensaverAuditor.Models;
using ScreensaverAuditor.Services;
using ScreensaverAuditor.Utils;

namespace ScreensaverAuditor
{
    public class Program
    {
        private static readonly ExcelExporter excelExporter = new();
        private static readonly ScreensaverAuditorService auditorService = new();

        static int Main(string[] args)
        {
            try
            {

                // 관리자 권한 확인
                if (!ScreensaverAuditorService.IsAdministrator())
                {
                    // 관리자 권한이 아닐 경우 경고 메시지 출력
                    Console.WriteLine("경고: 이 프로그램은 관리자 권한으로 실행해야 합니다.");
                    Console.WriteLine("관리자 권한으로 다시 실행해주세요.");
                    Console.WriteLine("\n종료하려면 아무 키나 누르세요...");
                    Console.ReadKey();
                    return 1;
                }

                // 명령줄 인수가 있는 경우 기존 방식으로 처리
                if (args.Length > 0)
                {
                    return ProcessCommandLineArguments(args);
                }

                // 명령줄 인수가 없는 경우 대화형 모드로 실행
                return RunInteractiveMode();
            }
            catch (Exception ex)
            {
                ConsoleHelper.HandleException(ex, args.Length == 0);
                return ConsoleHelper.GetErrorCode(ex);
            }
        }

        private static int RunInteractiveMode()
        {
            Console.WriteLine("화면 보호기 감사 도구 시작...");
            Console.WriteLine("감사 정책 상태 확인 중...");

            // 감사 정책 상태 확인
            bool isEnabled = auditorService.IsAuditPolicyEnabled();

            if (isEnabled)
            {
                Console.WriteLine("감사 정책이 활성화되어 있습니다.");
            }
            else
            {
                Console.WriteLine("감사 정책이 비활성되어 있습니다. 활성화를 시도합니다.");
                try
                {
                    auditorService.EnableAuditPolicy();
                    Console.WriteLine("감사 정책이 활성화되었습니다.");

                }
                catch (UnauthorizedAccessException)
                {
                    Console.WriteLine("경고: 관리자 권한이 없어 감사 정책을 활성화할 수 없습니다.");
                    Console.WriteLine("관리자 권한으로 프로그램을 다시 실행해주세요.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"감사 정책 활성화 중 오류 발생: {ex.Message}");
                }
            }

            // 도움말 표시
            Console.WriteLine("\n사용 가능한 명령어:");
            DisplayHelp();

            // 명령어 대기 루프
            while (true)
            {
                Console.Write("\n명령어 입력> ");
                string? input = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(input))
                    continue;

                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("quit", StringComparison.OrdinalIgnoreCase))
                {
                    return 0;
                }

                try
                {
                    // 입력된 명령어를 파싱하여 실행
                    string[] cmdArgs = SplitCommandLine(input);
                    ProcessCommandLineArguments(cmdArgs);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"명령 실행 중 오류: {ex.Message}");
                }
            }
        }

        // 인용부호, Escape 문자를 지원하도록 개선
        private static string[] SplitCommandLine(string commandLine)
        {
            var args = new List<string>();
            bool inQuotes = false;
            var current = new System.Text.StringBuilder();

            foreach (char c in commandLine)
            {
                if (c == '\"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (char.IsWhiteSpace(c) && !inQuotes)
                {
                    if (current.Length > 0)
                    {
                        args.Add(current.ToString());
                        current.Clear();
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            if (current.Length > 0) args.Add(current.ToString());
            return args.ToArray();
        }

        private static void DisplayHelp()
        {
            Console.WriteLine("  --help               - 이 도움말을 표시합니다");
            Console.WriteLine("  --start-date <날짜>      - 조회 시작 날짜 (yyyy-MM-dd 형식)");
            Console.WriteLine("  --end-date <날짜>        - 조회 종료 날짜 (yyyy-MM-dd 형식)");
            Console.WriteLine("  --user <사용자명>    - 특정 사용자만 조회");
            Console.WriteLine("  --output <파일명>   - 결과 저장 파일 경로");
            Console.WriteLine("  --enable-policy    - 감사 정책 활성화");
            Console.WriteLine("  exit, quit         - 프로그램 종료");
            Console.WriteLine("\n예시: --start-date 2023-01-01 --end-date 2025-01-31 --user Administrator");
        }

        private static int ProcessCommandLineArguments(string[] args)
        {
            try
            {
                var options = CommandLineParser.ParseCommandLineOptions(args);

                if (options.ShowHelp ||
                    (args.Length == 1 && (args[0].Equals("help", StringComparison.OrdinalIgnoreCase) ||
                                          args[0].Equals("--help", StringComparison.OrdinalIgnoreCase) ||
                                          args[0].Equals("-h", StringComparison.OrdinalIgnoreCase))))
                {
                    DisplayHelp();
                    return 0;
                }

                if (options.EnablePolicy)
                {
                    auditorService.EnableAuditPolicy();
                }

                var events = auditorService.GetScreensaverEvents(options.StartDate, options.EndDate, options.Username);
                var analysis = auditorService.AnalyzeScreensaverUsage(events);

                ConsoleHelper.DisplayResults(analysis);
                string dateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                SaveEventDetails(events, options.OutputPath ?? $"ScreensaverAuditResults_{dateTime}.xlsx");

                return 0;
            }
            catch (Exception ex)
            {
                ConsoleHelper.HandleException(ex, false);
                return ConsoleHelper.GetErrorCode(ex);
            }
        }

        private static void SaveEventDetails(List<ScreensaverEvent> events, string outputPath)
        {
            excelExporter.SaveToExcel(outputPath, events);
            Console.WriteLine($"결과가 '{outputPath}'에 저장되었습니다.");
        }
    }
}
