using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;                    // EPPlus
using OfficeOpenXml.Style;

namespace ORSResearchTool
{
    public static class ExcelHandler
    {
        // ── Read ORS IDs from input file ─────────────────────────────────────────────

        /// <summary>
        /// Reads the "ORS ID" column from the first worksheet of the supplied file.
        /// Skips blank/null cells.
        /// </summary>
        public static List<string> ReadOrsIds(string filePath)
        {
            var ids = new List<string>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var ws = package.Workbook.Worksheets[0];

            if (ws.Dimension == null)
                throw new InvalidOperationException("The selected Excel file appears to be empty.");

            // Find the header row containing "ORS ID" (case-insensitive)
            int headerRow    = -1;
            int orsIdColumn  = -1;

            for (int row = 1; row <= Math.Min(10, ws.Dimension.End.Row); row++)
            {
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    string cell = ws.Cells[row, col].GetValue<string>() ?? string.Empty;
                    if (cell.Trim().Equals("ORS ID", StringComparison.OrdinalIgnoreCase))
                    {
                        headerRow   = row;
                        orsIdColumn = col;
                        break;
                    }
                }
                if (headerRow != -1) break;
            }

            if (headerRow == -1)
                throw new InvalidOperationException("Column 'ORS ID' not found in the first 10 rows of the file.");

            for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
            {
                string id = ws.Cells[row, orsIdColumn].GetValue<string>();
                if (!string.IsNullOrWhiteSpace(id))
                    ids.Add(id.Trim());
            }

            return ids;
        }

        // ── Write output file ─────────────────────────────────────────────────────────

        public static void WriteResults(string outputPath, IList<OrsRecord> records)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("ORS Results");

            // ── Header row ────────────────────────────────────────────────────────────
            string[] headers = { "ORS ID", "Sent Date", "Received Date", "Difference (Days)", "Status" };

            for (int col = 1; col <= headers.Length; col++)
            {
                var cell = ws.Cells[1, col];
                cell.Value = headers[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.Name = "Arial";
                cell.Style.Font.Size = 10;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(68, 114, 196));
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            // ── Data rows ─────────────────────────────────────────────────────────────
            string dateFmt = "MM/dd/yyyy";

            for (int i = 0; i < records.Count; i++)
            {
                int row    = i + 2;
                var record = records[i];

                ws.Cells[row, 1].Value = record.OrsId;
                ws.Cells[row, 2].Value = record.SentDate.HasValue
                    ? record.SentDate.Value.ToString(dateFmt) : "NA";
                ws.Cells[row, 3].Value = record.ReceivedDate.HasValue
                    ? record.ReceivedDate.Value.ToString(dateFmt) : "NA";
                ws.Cells[row, 4].Value = record.Difference;
                ws.Cells[row, 5].Value = record.Status;

                // Colour-code by status
                var fill = ws.Cells[row, 1, row, 5];
                fill.Style.Font.Name = "Arial";
                fill.Style.Font.Size = 10;
                fill.Style.Fill.PatternType = ExcelFillStyle.Solid;

                switch (record.Status)
                {
                    case "Found":
                        fill.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(226, 239, 218));
                        break;
                    case "NA":
                        fill.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 242, 204));
                        break;
                    case "Error":
                        fill.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 213, 213));
                        break;
                    default:
                        fill.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                        break;
                }

                // Light border
                for (int col = 1; col <= 5; col++)
                    ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Hair);
            }

            // ── Auto-fit columns ──────────────────────────────────────────────────────
            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            // ── Summary row ───────────────────────────────────────────────────────────
            int summaryRow = records.Count + 3;
            ws.Cells[summaryRow, 1].Value = "Total Records";
            ws.Cells[summaryRow, 2].Value = records.Count;
            ws.Cells[summaryRow + 1, 1].Value = "Found";
            ws.Cells[summaryRow + 1, 2].Value = CountByStatus(records, "Found");
            ws.Cells[summaryRow + 2, 1].Value = "Not Found (NA)";
            ws.Cells[summaryRow + 2, 2].Value = CountByStatus(records, "NA");
            ws.Cells[summaryRow + 3, 1].Value = "Errors";
            ws.Cells[summaryRow + 3, 2].Value = CountByStatus(records, "Error");

            for (int r = summaryRow; r <= summaryRow + 3; r++)
            {
                ws.Cells[r, 1].Style.Font.Bold = true;
                ws.Cells[r, 1].Style.Font.Name = "Arial";
                ws.Cells[r, 2].Style.Font.Name = "Arial";
            }

            package.SaveAs(new FileInfo(outputPath));
        }

        private static int CountByStatus(IList<OrsRecord> records, string status)
        {
            int count = 0;
            foreach (var r in records)
                if (r.Status == status) count++;
            return count;
        }
    }
}
