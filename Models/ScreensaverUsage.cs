// Models/ScreensaverUsage.cs
using System;
using System.Collections.Generic;

namespace ScreensaverAuditor.Models
{
    public class ScreensaverUsage
    {
        public string Username { get; }
        public string ComputerName { get; set; } = "";
        public int ScreensaverActivationCount { get; set; }
        public TimeSpan TotalScreensaverDuration { get; set; }
        public List<(DateTime Start, DateTime? End, int EventId)> ScreensaverPeriods { get; }

        public ScreensaverUsage(string user)
        {
            Username = user;
            ScreensaverPeriods = new List<(DateTime, DateTime?, int)>();
        }
    }
}
