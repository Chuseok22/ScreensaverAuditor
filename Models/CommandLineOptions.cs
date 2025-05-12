// Models/CommandLineOptions.cs
using System;

namespace ScreensaverAuditor.Models
{
    public class CommandLineOptions
    {
        public bool EnablePolicy { get; set; }
        public bool ShowHelp { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string? Username { get; set; }
        public string? OutputPath { get; set; }
    }
}
