using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ScreensaverAuditor.Models;
using ScreensaverAuditor.Services;
using ScreensaverAuditor.Utils;
using ScreensaverAuditor.UI;

namespace ScreensaverAuditor
{
    public class Program
    {
        private static readonly ExcelExporter excelExporter = new();
        private static readonly ScreensaverAuditorService auditorService = new();
        
        // GUI 모드인지 여부
        private static bool _guiMode = true;
        
        // 콘솔 창을 숨기기 위한 Win32 API
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        [STAThread]
        static int Main(string[] args)
        {
            try
            {
                // 명령줄 옵션 파악
                if (args.Length > 0)
                {
                    // "--console" 인수가 있으면 콘솔 모드로 실행
                    if (Array.IndexOf(args, "--console") >= 0)
                    {
                        _guiMode = false;
                    }
                    else
                    {
                        // GUI 모드에서는 콘솔 창 숨기기
                        HideConsoleWindow();
                    }
                }
                else
                {
                    // 인수가 없으면 GUI 모드로 실행하고 콘솔 창 숨기기
                    HideConsoleWindow();
                }
                
                // 관리자 권한 확인 (GUI 모드, 콘솔 모드 모두)
                bool isAdmin = ScreensaverAuditorService.IsAdministrator();

                if (_guiMode)
                {
                    // GUI 모드로 실행
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    
                    if (!isAdmin)
                    {
                        // GUI 모드에서 관리자 권한이 없을 때 메시지 표시
                        MessageBox.Show(
                            "이 프로그램은 Windows 이벤트 로그 및 감사 정책에 접근하기 위해 관리자 권한이 필요합니다.\n\n" +
                            "관리자 권한으로 다시 실행해주세요.",
                            "관리자 권한 필요",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return 1;
                    }
                    
                    Application.Run(new MainForm());
                    return 0;
                }
                else
                {
                    // 콘솔 모드에서 관리자 권한 확인
                    if (!isAdmin)
                    {
                        Console.WriteLine("경고: 이 프로그램은 관리자 권한으로 실행해야 합니다.");
                        Console.WriteLine("관리자 권한으로 다시 실행해주세요.");
                        Console.WriteLine("\n종료하려면 아무 키나 누르세요...");
                        Console.ReadKey();
                        return 1;
                    }
                    
                    // 명령줄 인수가 있으면 명령줄 모드로 실행
                    if (args.Length > 0)
                    {
                        return ProcessCommandLineArguments(args);
                    }
                    
                    // 명령줄 인수가 없으면 대화형 모드로 실행
                    return RunInteractiveMode();
                }
            }
            catch (Exception ex)
            {
                if (!_guiMode)
                {
                    ConsoleHelper.HandleException(ex, false);
                }
                else
                {
                    MessageBox.Show($"오류가 발생했습니다: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return ConsoleHelper.GetErrorCode(ex);
            }
        }

        // 콘솔 창 숨기기 메서드
        private static void HideConsoleWindow()
        {
            var handle = GetConsoleWindow();
            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_HIDE);
            }
        }

        // 대화형 모드 실행
        private static int RunInteractiveMode()
        {
            Console.WriteLine("화면 보호기 감사 도구 시작...");
            Console.WriteLine("감사 정책 상태 확인 중...");

            // 감사 정책 상태 확인
            try
            {
                bool isEnabled = auditorService.IsAuditPolicyEnabled();
                if (isEnabled)
                {
                    Console.WriteLine("감사 정책이 활성화되어 있습니다.");
                }
                else
                {
                    Console.WriteLine("감사 정책이 비활성화되어 있습니다.");
                    Console.WriteLine("감사 정책을 활성화하려면 --enable-policy 옵션을 사용하세요.");
                }
            }
            catch (Exception)
            {
                Console.WriteLine("감사 정책 상태를 확인할 수 없습니다. 관리자 권한으로 실행하세요.");
            }

            // 대화형 명령 처리
            bool exitRequested = false;
            while (!exitRequested)
            {
                Console.WriteLine("\n사용 가능한 명령어:");
                Console.WriteLine("  help        - 사용 방법 표시");
                Console.WriteLine("  audit       - 기본 설정으로 감사 실행");
                Console.WriteLine("  enable      - 감사 정책 활성화");
                Console.WriteLine("  exit, quit  - 프로그램 종료");
                Console.Write("\n> ");

                string? input = Console.ReadLine()?.Trim().ToLower();
                switch (input)
                {
                    case "help":
                        DisplayHelp();
                        break;
                        
                    case "audit":
                        RunDefaultAudit();
                        break;
                        
                    case "enable":
                        EnableAuditPolicy();
                        break;
                        
                    case "exit":
                    case "quit":
                        exitRequested = true;
                        break;
                        
                    default:
                        Console.WriteLine("알 수 없는 명령입니다. 'help'를 입력하여 도움말을 확인하세요.");
                        break;
                }
            }

            return 0;
        }

        // 기본 감사 실행
        private static void RunDefaultAudit()
        {
            try
            {
                var options = new CommandLineOptions
                {
                    RunAudit = true,
                    StartDate = DateTime.Now.AddDays(-7),
                    EndDate = DateTime.Now,
                };

                Console.WriteLine($"기간: {options.StartDate:yyyy-MM-dd} ~ {options.EndDate:yyyy-MM-dd}");
                
                var events = auditorService.GetScreensaverEvents(options.StartDate, options.EndDate, options.Username);
                
                if (events == null || events.Count == 0)
                {
                    Console.WriteLine("\n[결과 없음] 지정한 기간에 화면보호기 이벤트가 없습니다.");
                    Console.WriteLine($"검색 기간: {options.StartDate:yyyy-MM-dd} ~ {options.EndDate:yyyy-MM-dd}");
                    
                    if (!string.IsNullOrEmpty(options.Username))
                    {
                        Console.WriteLine($"사용자: {options.Username}");
                    }
                    
                    Console.WriteLine("\n다른 검색 조건을 시도해 보세요.");
                    return;
                }
                
                string outputPath = "화면보호기감사_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                excelExporter.SaveToExcel(outputPath, events);
                
                Console.WriteLine($"\n총 {events.Count}개 이벤트를 찾았습니다.");
                Console.WriteLine($"결과가 '{outputPath}'에 저장되었습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"감사 실행 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        // 감사 정책 활성화
        private static void EnableAuditPolicy()
        {
            try
            {
                auditorService.EnableAuditPolicy();
                Console.WriteLine("감사 정책이 활성화되었습니다.");
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("감사 정책을 활성화하려면 관리자 권한이 필요합니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"정책 활성화 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        // 명령줄 인수 처리
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
                
                // --audit 명령어 처리 - 기본값을 그대로 사용하므로 별도 처리 필요 없음
                if (options.RunAudit)
                {
                    Console.WriteLine("기본 설정으로 감사를 실행합니다 (7일 전 ~ 오늘)");
                    // StartDate, EndDate는 이미 기본값으로 설정되어 있어 별도 설정 불필요
                }

                var events = auditorService.GetScreensaverEvents(options.StartDate, options.EndDate, options.Username);

                if (events == null || events.Count == 0)
                {
                    Console.WriteLine("\n[결과 없음] 지정한 기간에 화면보호기 이벤트가 없습니다.");
                    Console.WriteLine($"검색 기간: {options.StartDate:yyyy-MM-dd} ~ {options.EndDate:yyyy-MM-dd}");
                    
                    if (!string.IsNullOrEmpty(options.Username))
                    {
                        Console.WriteLine($"사용자: {options.Username}");
                    }
                    
                    Console.WriteLine("\n다른 검색 조건을 시도해 보세요.");
                    return 0;
                }

                string outputPath = options.OutputPath ?? "화면보호기감사_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                excelExporter.SaveToExcel(outputPath, events);

                Console.WriteLine($"총 {events.Count}개의 이벤트를 찾았습니다.");
                Console.WriteLine($"결과가 '{outputPath}'에 저장되었습니다.");

                return 0;
            }
            catch (Exception ex)
            {
                ConsoleHelper.HandleException(ex, false);
                return ConsoleHelper.GetErrorCode(ex);
            }
        }

        private static void DisplayHelp()
        {
            Console.WriteLine("화면보호기 감사 도구 - 실행 모드");
            Console.WriteLine("----------------------------");
            Console.WriteLine("  기본 실행               - 그래픽 사용자 인터페이스(GUI) 모드로 실행");
            Console.WriteLine("  --console             - 콘솔 모드로 실행 (명령줄 인터페이스)");
            Console.WriteLine("\n명령줄 옵션:");
            Console.WriteLine("  --help               - 이 도움말을 표시합니다");
            Console.WriteLine("  --audit              - 기본 설정으로 감사 실행 (7일 전 ~ 오늘)");
            Console.WriteLine("  --start-date <날짜>      - 조회 시작 날짜 (yyyy-MM-dd 형식)");
            Console.WriteLine("  --end-date <날짜>        - 조회 종료 날짜 (yyyy-MM-dd 형식)");
            Console.WriteLine("  --user <사용자명>    - 특정 사용자만 조회");
            Console.WriteLine("  --output <파일명>   - 결과 저장 파일 경로");
            Console.WriteLine("  --enable-policy    - 감사 정책 활성화");
            Console.WriteLine("  exit, quit         - 프로그램 종료");
            Console.WriteLine("\n예시: --console --start-date 2023-01-01 --end-date 2025-01-31 --user Administrator");
            Console.WriteLine("\n간단 실행: --audit");
            Console.WriteLine("\nGUI 모드에서는 모든 기능을 버튼과 폼을 통해 쉽게 사용할 수 있습니다.");
        }
    }
}
