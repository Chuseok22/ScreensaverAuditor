// ViewModels/ScreensaverAuditorViewModel.cs
using System;
using System.Collections.Generic;
using ScreensaverAuditor.Models;
using ScreensaverAuditor.Services;

namespace ScreensaverAuditor.ViewModels
{
    // 상태 변경 이벤트 인자
    public class StatusChangedEventArgs : EventArgs
    {
        public string Status { get; }

        public StatusChangedEventArgs(string status)
        {
            Status = status;
        }
    }

    // 감사 완료 이벤트 인자
    public class AuditCompletedEventArgs : EventArgs
    {
        public List<ScreensaverEvent> Events { get; }
        public string OutputPath { get; }

        public AuditCompletedEventArgs(List<ScreensaverEvent> events, string outputPath)
        {
            Events = events;
            OutputPath = outputPath;
        }
    }

    public class ScreensaverAuditorViewModel
    {
        private readonly ScreensaverAuditorService _auditorService;
        private readonly ExcelExporter _excelExporter;

        // 이벤트 선언
        public event EventHandler<StatusChangedEventArgs> StatusChanged;
        public event EventHandler<AuditCompletedEventArgs> AuditCompleted;

        public ScreensaverAuditorViewModel()
        {
            _auditorService = new ScreensaverAuditorService();
            _excelExporter = new ExcelExporter();
        }

        // 상태 변경 이벤트 발생 메서드
        protected virtual void OnStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, new StatusChangedEventArgs(status));
        }

        // 감사 완료 이벤트 발생 메서드
        protected virtual void OnAuditCompleted(List<ScreensaverEvent> events, string outputPath)
        {
            AuditCompleted?.Invoke(this, new AuditCompletedEventArgs(events, outputPath));
        }

        // 감사 실행 메서드
        public void RunAudit(CommandLineOptions options)
        {
            try
            {
                OnStatusChanged("감사 정보 수집 중...");

                // 출력 경로 설정
                string outputPath = string.IsNullOrEmpty(options.OutputPath) ? 
                    "화면보호기감사.xlsx" : options.OutputPath;
                
                // 이벤트 수집
                OnStatusChanged("화면보호기 이벤트 수집 중...");
                var events = _auditorService.GetScreensaverEvents(
                    options.StartDate, options.EndDate, options.Username);

                if (events == null || events.Count == 0)
                {
                    OnStatusChanged("이벤트 없음: 결과가 없습니다.");
                    OnAuditCompleted(new List<ScreensaverEvent>(), null);
                    return;
                }

                // Excel로 내보내기
                OnStatusChanged($"{events.Count}개의 이벤트 발견. Excel 파일 생성 중...");
                _excelExporter.SaveToExcel(outputPath, events);
                
                OnStatusChanged($"감사 완료: {events.Count}개의 이벤트가 {outputPath}에 저장되었습니다.");
                OnAuditCompleted(events, outputPath);
            }
            catch (Exception ex)
            {
                OnStatusChanged($"오류 발생: {ex.Message}");
                throw;
            }
        }

        // 감사 정책 활성화 메서드
        public void EnableAuditPolicy()
        {
            try
            {
                OnStatusChanged("감사 정책 활성화 중...");
                _auditorService.EnableAuditPolicy();
                OnStatusChanged("감사 정책이 성공적으로 활성화되었습니다.");
            }
            catch (Exception)
            {
                OnStatusChanged("감사 정책 활성화 실패!");
                throw;
            }
        }
    }
}
