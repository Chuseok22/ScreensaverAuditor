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
        private void EnsureValidPath(string outputPath)
        {
            string finalPath = !outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? outputPath + ".xlsx"
                : outputPath;

            if (File.Exists(finalPath))
            {
                try { File.Delete(finalPath); }
                catch (IOException) { throw new IOException("Excel 파일이 열려 있습니다. 닫고 다시 시도하세요."); }
            }

            var dir = Path.GetDirectoryName(finalPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

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

        private void AddData(IXLWorksheet worksheet, List<ScreensaverEvent> events)
        {
            int row = 2;
            foreach (var evt in events)
            {
                AddEventRow(worksheet, row++, evt);
            }
        }

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

        private void ApplyFormatting(IXLWorksheet worksheet)
        {
            var headerRange = worksheet.Range(1, 1, 1, 10);  // 10개 열
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            worksheet.SheetView.FreezeRows(1);
            worksheet.Columns().AdjustToContents();
        }

        public void SaveToExcel(string outputPath, List<ScreensaverEvent> events)
        {
            EnsureValidPath(outputPath);

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Events");

            AddHeaders(worksheet);
            AddData(worksheet, events);
            ApplyFormatting(worksheet);

            workbook.SaveAs(outputPath);
        }
    }
}
