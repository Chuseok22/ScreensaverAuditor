// Models/ScreensaverEvent.cs
using System;

namespace ScreensaverAuditor.Models
{
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
}
