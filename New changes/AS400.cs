// ── Feature 1: Read ICN from (row 4, col 44, len 10) ─────────────────────────
/// <summary>
/// Reads the ICN after OpenOrs() succeeds.
/// ICN format: 2 alpha chars + 8 numeric chars (e.g. FP58697495).
/// Returns null if the field is blank or doesn't match the pattern.
/// </summary>
public string ReadIcn()
{
    string raw = GetText(4, 44, 10, trim: true);
    if (string.IsNullOrWhiteSpace(raw)) return null;

    // Validate: first 2 chars alpha, remaining 8 numeric
    if (raw.Length == 10
        && char.IsLetter(raw[0]) && char.IsLetter(raw[1])
        && long.TryParse(raw.Substring(2), out _))
        return raw;

    return raw; // return as-is if format is unexpected — let caller decide
}

// ── Feature 2: Capture blue comment line ─────────────────────────────────────
/// <summary>
/// After a matching 224/007 (or 007/224) line is found on the current screen
/// at <paramref name="matchRow"/>, reads consecutive rows below it that are
/// coloured blue (mdlUnet.getcolorscheme() == 2) and returns their text joined
/// with a space.  White rows (== 1) terminate the scan.
/// </summary>
public string ReadCommentAfterRow(int matchRow)
{
    var sb = new System.Text.StringBuilder();

    for (int row = matchRow + 1; row <= SCREEN_ROWS; row++)
    {
        int colorScheme = mdlUnet.getcolorscheme(row, 1);
        if (colorScheme != 2) break;   // blue == 2; white == 1 → stop

        string line = GetText(row, 1, SCREEN_COLS, trim: true);
        if (!string.IsNullOrEmpty(line))
        {
            if (sb.Length > 0) sb.Append(" ");
            sb.Append(line);
        }
    }

    return sb.Length > 0 ? sb.ToString() : null;
}




// Replace the existing ScanCurrentScreen method in AS400Session.cs

private string ScanCurrentScreen(string targetId1, string targetId2, out int matchedRow)
{
    matchedRow = -1;
    for (int row = 1; row <= SCREEN_ROWS; row++)
    {
        string id1 = GetText(row, OFFID1_COL, 3, trim: true);
        string id2 = GetText(row, OFFID2_COL, 3, trim: true);

        if (id1 == targetId1 && id2 == targetId2)
        {
            matchedRow = row;
            return GetText(row, DATE_COL, 10, trim: true);
        }
    }
    return null;
}



// Replace the existing ScanPages method in AS400Session.cs

private string ScanPages(string targetId1, string targetId2, out int matchedRow)
{
    matchedRow = -1;
    const int MAX_PAGES = 500;

    for (int page = 0; page < MAX_PAGES; page++)
    {
        string hit = ScanCurrentScreen(targetId1, targetId2, out matchedRow);
        if (hit != null) return hit;

        if (GetText(1, 1, SCREEN_COLS, trim: true).Contains(LAST_PAGE))
            return null;

        SendF8();
    }

    return null;
}



// Replace both existing Find methods in AS400Session.cs

public string FindSentDate(out string comment)
{
    string date = ScanPages(targetId1: "224", targetId2: "007", out int row);
    comment = row > 0 ? ReadCommentAfterRow(row) : null;
    return date;
}

public string FindReceivedDate(out string comment)
{
    SendF7();
    string date = ScanPages(targetId1: "007", targetId2: "224", out int row);
    comment = row > 0 ? ReadCommentAfterRow(row) : null;
    return date;
}


