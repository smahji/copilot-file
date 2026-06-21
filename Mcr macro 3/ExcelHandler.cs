using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ORSResearchTool
{
    public static class ExcelHandler
    {
        // ── Column layout — single source of truth for all methods ────────────────────
        // Order must match WriteDataRow column assignments below.
        private static readonly string[] Headers =
        {
            "ORS ID",           // col 1
            "ICN",              // col 2
            "Sent Date",        // col 3
            "Sent Comment",     // col 4
            "Received Date",    // col 5
            "Recv Comment",     // col 6
            "Difference (Days)",// col 7
            "Remark Codes",     // col 8
            "Adj ID",           // col 9
            "Status"            // col 10
        };

        private const string SHEET_NAME = "ORS Results";
        private const string DATE_FMT   = "MM/dd/yyyy";

        // ── Read ORS IDs from input file ──────────────────────────────────────────────

        /// <summary>
        /// Reads the "ORS ID" column from the first worksheet.
        /// Searches the first 10 rows for the header (case-insensitive).
        /// Skips blank/null cells in the data rows.
        /// </summary>
        public static List<string> ReadOrsIds(string filePath)
        {
            var ids = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets[0];

                if (ws.Dimension == null)
                    throw new InvalidOperationException(
                        "The selected Excel file appears to be empty.");

                int headerRow = -1, orsIdCol = -1;

                for (int row = 1; row <= Math.Min(10, ws.Dimension.End.Row); row++)
                {
                    for (int col = 1; col <= ws.Dimension.End.Column; col++)
                    {
                        string cell = ws.Cells[row, col].GetValue<string>() ?? string.Empty;
                        if (cell.Trim().Equals("ORS ID", StringComparison.OrdinalIgnoreCase))
                        {
                            headerRow = row;
                            orsIdCol  = col;
                            break;
                        }
                    }
                    if (headerRow != -1) break;
                }

                if (headerRow == -1)
                    throw new InvalidOperationException(
                        "Column 'ORS ID' not found in the first 10 rows of the input file.");

                for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
                {
                    string id = ws.Cells[row, orsIdCol].GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(id))
                        ids.Add(id.Trim());
                }
            }

            return ids;
        }

        // ── Checkpoint: create output file with header row before loop starts ─────────

        /// <summary>
        /// Creates the output file and writes the styled header row.
        /// Called ONCE by MainForm before processing begins.
        /// Deletes any pre-existing file at the same path so we always start clean.
        /// </summary>
        public static void InitialiseOutputFile(string outputPath)
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add(SHEET_NAME);
                WriteHeaderRow(ws);
                package.SaveAs(new FileInfo(outputPath));
            }
        }

        // ── Checkpoint: append one completed record immediately after each ORS ─────────

        /// <summary>
        /// Opens the existing output file, appends <paramref name="record"/> as the
        /// next data row, and saves immediately.
        /// Called after EVERY ORS — Found, NA, or Error — so the file always
        /// reflects exactly how far processing has reached.
        /// Thread-safe via a per-file lock.
        /// </summary>
        public static void AppendResult(string outputPath, OrsRecord record)
        {
            lock (GetFileLock(outputPath))
            {
                using (var package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    var ws = package.Workbook.Worksheets[SHEET_NAME]
                             ?? package.Workbook.Worksheets.Add(SHEET_NAME);

                    // Safety: if sheet lost its header (shouldn't happen), recreate it
                    if (ws.Dimension == null)
                        WriteHeaderRow(ws);

                    int nextRow = ws.Dimension.End.Row + 1;
                    WriteDataRow(ws, nextRow, record);
                    package.Save();
                }
            }
        }

        // ── Finalise: append summary block + auto-fit on completion or cancel ─────────

        /// <summary>
        /// Opens the checkpoint file and appends the summary block at the bottom,
        /// then auto-fits all columns.
        /// Safe to call after a cancellation — data rows are already intact.
        /// </summary>
        public static void FinaliseOutputFile(string outputPath, IList<OrsRecord> allRecords)
        {
            if (!File.Exists(outputPath)) return;

            lock (GetFileLock(outputPath))
            {
                using (var package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    var ws = package.Workbook.Worksheets[SHEET_NAME];
                    if (ws == null) return;

                    // Auto-fit all columns
                    if (ws.Dimension != null)
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // Summary block — 2 blank rows below last data row
                    int summaryRow = (ws.Dimension?.End.Row ?? 1) + 2;

                    WriteSummaryRow(ws, summaryRow,     "Total Records",  allRecords.Count);
                    WriteSummaryRow(ws, summaryRow + 1, "Found",          CountByStatus(allRecords, "Found"));
                    WriteSummaryRow(ws, summaryRow + 2, "Not Found (NA)", CountByStatus(allRecords, "NA"));
                    WriteSummaryRow(ws, summaryRow + 3, "Errors",         CountByStatus(allRecords, "Error"));

                    package.Save();
                }
            }
        }

        // ── WriteResults: convenience method for bulk in-memory write ────────────────
        // Kept for compatibility; internally uses the checkpoint methods.

        public static void WriteResults(string outputPath, IList<OrsRecord> records)
        {
            InitialiseOutputFile(outputPath);
            foreach (var r in records)
                AppendResult(outputPath, r);
            FinaliseOutputFile(outputPath, records);
        }

        // ── Private: header row ───────────────────────────────────────────────────────

        private static void WriteHeaderRow(ExcelWorksheet ws)
        {
            for (int col = 1; col <= Headers.Length; col++)
            {
                var cell = ws.Cells[1, col];
                cell.Value = Headers[col - 1];
                cell.Style.Font.Bold      = true;
                cell.Style.Font.Name      = "Arial";
                cell.Style.Font.Size      = 10;
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.HorizontalAlignment          = ExcelHorizontalAlignment.Center;
                cell.Style.Fill.PatternType             = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        // ── Private: data row ─────────────────────────────────────────────────────────

        private static void WriteDataRow(ExcelWorksheet ws, int row, OrsRecord record)
        {
            // Col 1 — ORS ID
            ws.Cells[row, 1].Value = record.OrsId;

            // Col 2 — ICN
            ws.Cells[row, 2].Value = record.Icn ?? "NA";

            // Col 3 — Sent Date
            ws.Cells[row, 3].Value = record.SentDate.HasValue
                ? record.SentDate.Value.ToString(DATE_FMT) : "NA";

            // Col 4 — Sent Comment (blue line after 224/007)
            ws.Cells[row, 4].Value = record.SentComment ?? "NA";

            // Col 5 — Received Date
            ws.Cells[row, 5].Value = record.ReceivedDate.HasValue
                ? record.ReceivedDate.Value.ToString(DATE_FMT) : "NA";

            // Col 6 — Recv Comment (blue line after 007/224)
            ws.Cells[row, 6].Value = record.RecvComment ?? "NA";

            // Col 7 — Difference (TAT days)
            ws.Cells[row, 7].Value = record.Difference ?? "NA";

            // Col 8 — Remark Codes (comma-separated from MHI Session B)
            ws.Cells[row, 8].Value = record.RemarkCodes ?? "NA";

            // Col 9 — Adjuster ID (last page of MHI)
            ws.Cells[row, 9].Value = record.AdjusterId ?? "NA";

            // Col 10 — Status
            ws.Cells[row, 10].Value = record.Status;

            // ── Row colour by status ──────────────────────────────────────────
            Color rowColour;
            switch (record.Status)
            {
                case "Found": rowColour = Color.FromArgb(226, 239, 218); break; // light green
                case "NA":    rowColour = Color.FromArgb(255, 242, 204); break; // light yellow
                case "Error": rowColour = Color.FromArgb(255, 213, 213); break; // light red
                default:      rowColour = Color.White;                   break;
            }

            var range = ws.Cells[row, 1, row, Headers.Length];
            range.Style.Font.Name        = "Arial";
            range.Style.Font.Size        = 10;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(rowColour);

            for (int col = 1; col <= Headers.Length; col++)
                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Hair);
        }

        // ── Private: summary row ──────────────────────────────────────────────────────

        private static void WriteSummaryRow(ExcelWorksheet ws, int row, string label, int value)
        {
            ws.Cells[row, 1].Value           = label;
            ws.Cells[row, 2].Value           = value;
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Style.Font.Name = "Arial";
            ws.Cells[row, 2].Style.Font.Name = "Arial";
        }

        // ── Private: status counter ───────────────────────────────────────────────────

        private static int CountByStatus(IList<OrsRecord> records, string status)
        {
            int n = 0;
            foreach (var r in records)
                if (r.Status == status) n++;
            return n;
        }

        // ── Private: per-file lock registry ──────────────────────────────────────────
        // Ensures AppendResult/FinaliseOutputFile never write concurrently to the same
        // file if called from multiple threads (future-proofing).

        private static readonly ConcurrentDictionary<string, object> _fileLocks =
            new ConcurrentDictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        private static object GetFileLock(string path)
            => _fileLocks.GetOrAdd(Path.GetFullPath(path), _ => new object());
    }
}
