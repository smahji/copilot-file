// In ExcelHandler.cs — replace the Headers line

private static readonly string[] Headers =
{
    "ORS ID", "ICN", "Sent Date", "Sent Comment",
    "Received Date", "Recv Comment",
    "Difference (Days)", "Remark Codes", "Adj ID", "Status"
};



// In ExcelHandler.cs — replace WriteDataRow entirely

private static void WriteDataRow(ExcelWorksheet ws, int row, OrsRecord record)
{
    ws.Cells[row, 1].Value  = record.OrsId;
    ws.Cells[row, 2].Value  = record.Icn        ?? "NA";
    ws.Cells[row, 3].Value  = record.SentDate.HasValue
        ? record.SentDate.Value.ToString(DATE_FMT) : "NA";
    ws.Cells[row, 4].Value  = record.SentComment  ?? "NA";
    ws.Cells[row, 5].Value  = record.ReceivedDate.HasValue
        ? record.ReceivedDate.Value.ToString(DATE_FMT) : "NA";
    ws.Cells[row, 6].Value  = record.RecvComment  ?? "NA";
    ws.Cells[row, 7].Value  = record.Difference;
    ws.Cells[row, 8].Value  = record.RemarkCodes  ?? "NA";
    ws.Cells[row, 9].Value  = record.AdjusterId   ?? "NA";
    ws.Cells[row, 10].Value = record.Status;

    Color rowColour;
    switch (record.Status)
    {
        case "Found": rowColour = Color.FromArgb(226, 239, 218); break;
        case "NA":    rowColour = Color.FromArgb(255, 242, 204); break;
        case "Error": rowColour = Color.FromArgb(255, 213, 213); break;
        default:      rowColour = Color.White;                   break;
    }

    var range = ws.Cells[row, 1, row, Headers.Length];
    range.Style.Font.Name            = "Arial";
    range.Style.Font.Size            = 10;
    range.Style.Fill.PatternType     = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(rowColour);

    for (int col = 1; col <= Headers.Length; col++)
        ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Hair);
}