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
        
        // 추가 정보 필드
        public string EventType { get; }        // 화면보호기 시작/종료
        public string TaskDisplayName { get; }  // 이벤트 작업 이름
        public string ActivityId { get; }       // 활동 ID
        public string ProviderName { get; }     // 이벤트 제공자
        public string Keywords { get; }         // 감사 결과 (성공/실패 등)

        public ScreensaverEvent(
            DateTime ts,
            int id,
            string computer,
            string message,
            string username,
            string domain,
            string securityId,
            string logonId,
            string sessionId,
            string taskDisplayName = "",
            string activityId = "",
            string providerName = "",
            string keywords = "")
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
            
            // 이벤트 타입 설정 (4802: 시작, 4803: 종료)
            EventType = id == 4802 ? "화면보호기 시작" : "화면보호기 종료";
            
            // 추가 정보 설정
            TaskDisplayName = taskDisplayName;
            ActivityId = activityId;
            ProviderName = providerName;
            Keywords = keywords;
        }
    }
}
