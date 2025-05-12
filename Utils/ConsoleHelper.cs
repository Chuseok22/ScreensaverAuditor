// Utils/ConsoleHelper.cs
using System;
using ScreensaverAuditor.Models;
using System.Collections.Generic;

namespace ScreensaverAuditor.Utils
{
    public static class ConsoleHelper
    {
        public static void DisplayResults(List<ScreensaverUsage> analysis)
        {
            foreach (var usage in analysis)
            {
                Console.WriteLine($"사용자: {usage.Username}");
                Console.WriteLine($"화면보호기 실행 횟수: {usage.ScreensaverActivationCount}");
                Console.WriteLine($"총 지속 시간: {usage.TotalScreensaverDuration.TotalHours:F2} 시간");
                Console.WriteLine("───────────────────────────────────────");
            }
        }

        public static void WaitForKey()
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        public static void HandleException(Exception ex, bool waitForKey)
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

        public static int GetErrorCode(Exception ex) => ex switch
        {
            UnauthorizedAccessException => 1,
            ArgumentException => 2,
            _ => 3
        };
    }
}
