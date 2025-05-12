// Services/ExcelExporter.cs
using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ScreensaverAuditor.Models;

namespace ScreensaverAuditor.Services
{
    public class ExcelExporter
    {
        // 엑셀 파일을 저장합니다.
        public void SaveToExcel(string outputPath, List<ScreensaverEvent> events)
        {
            outputPath = EnsureValidPath(outputPath);

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Events");

            AddHeaders(worksheet);
            AddData(worksheet, events);
            ApplyFormatting(worksheet);

            workbook.SaveAs(outputPath);
        }

        // 엑셀 파일을 저장할 경로를 확인하고, 필요시 디렉토리를 생성합니다.
        private string EnsureValidPath(string outputPath)
        {
            string finalPath = !outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? outputPath + ".xlsx"
                : outputPath;

            // 이미 해당 경로에 파일이 존재하는 경우
            if (File.Exists(finalPath))
            {
                try { File.Delete(finalPath); } // 기존 파일 삭제
                catch (IOException) { throw new IOException("Excel 파일이 열려 있습니다. 닫고 다시 시도하세요."); }
            }

            var dir = Path.GetDirectoryName(finalPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return finalPath;
        }

        // 엑셀 파일에 헤더를 추가합니다.
        private void AddHeaders(IXLWorksheet worksheet)
        {
            string[] headers = {
                "시간",
                "이벤트 ID",
                "PC 관리번호",
                "메시지",
                "계정 이름",
                "계정 도메인",
                "보안 ID",
                "로그온 ID",
                "세션 ID",
                "지속시간(분)"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
            }
        }

        // 엑셀 파일에 데이터를 추가합니다.
        private void AddData(IXLWorksheet worksheet, List<ScreensaverEvent> events)
        {
            int row = 2;
            foreach (var evt in events)
            {
                AddEventRow(worksheet, row++, evt);
            }
        }

        // 각 이벤트에 대한 정보를 엑셀 파일에 추가합니다.
        private void AddEventRow(IXLWorksheet worksheet, int row, ScreensaverEvent evt)
        {
            // 기본 정보
            var dateCell = worksheet.Cell(row, 1);
            dateCell.Value = evt.Timestamp;
            dateCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";

            worksheet.Cell(row, 2).Value = evt.EventId;
            worksheet.Cell(row, 3).Value = evt.ComputerName;
            worksheet.Cell(row, 4).Value = evt.Message;
            worksheet.Cell(row, 5).Value = evt.Username;
            worksheet.Cell(row, 6).Value = evt.AccountDomain;
            worksheet.Cell(row, 7).Value = evt.SecurityId;
            worksheet.Cell(row, 8).Value = evt.LogonId;
            worksheet.Cell(row, 9).Value = evt.SessionId;

            // 지속시간 (분)
            var durationCell = worksheet.Cell(row, 10);
            if (evt.Duration.HasValue)
            {
                durationCell.Value = Math.Round(evt.Duration.Value.TotalMinutes, 2);
            }
            else
            {
                durationCell.Value = string.Empty;
            }
        }

        // 엑셀 파일의 서식을 설정합니다.
        private void ApplyFormatting(IXLWorksheet worksheet)
        {
            var headerRange = worksheet.Range(1, 1, 1, 10);  // 10개 열
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            worksheet.SheetView.FreezeRows(1);
            worksheet.Columns().AdjustToContents();
        }
    }
}
