using System;
using System.Threading;
using UnetHelper; // Reference to UnetHelper.dll

namespace ORSResearchTool
{
    /// <summary>
    /// Wraps the UnetHelper terminal session for AS400/111W navigation.
    /// All row/col coordinates follow 1-based terminal screen indexing.
    /// </summary>
    public class AS400Session : IDisposable
    {
        // ── UnetHelper session object (adjust type name to match your DLL's public API) ──
        private Session _session;

        private readonly string _host;
        private readonly string _paylocLogin;
        private readonly string _paylocPassword;

        // Configurable delays (ms) – increase on slow hosts
        public int ActionDelayMs { get; set; } = 500;
        public int PageScrollDelayMs { get; set; } = 300;

        // ── Coordinates ────────────────────────────────────────────────────────────────
        // FN field on ORS inquiry screen: row 22, col 5, length 2
        private const int FN_ROW = 22;
        private const int FN_COL = 5;

        // ORS# field: row 24, col 7
        private const int ORS_ROW = 24;
        private const int ORS_COL = 7;

        // Off-ID columns on the comment lines
        private const int OFFID1_COL = 9;   // First off-id  (expect "224")
        private const int OFFID2_COL = 25;  // Second off-id (expect "007")
        private const int DATE_COL   = 2;   // Date column

        // Sentinel string on last-page
        private const string LAST_PAGE_TEXT = "LAST PAGE OF NON-SYSTEM COMMENTS";

        // Main menu sentinel
        private const string MAIN_MENU_TEXT = "ONLINE ROUTING SYSTEM MAIN MENU";

        public AS400Session(string host, string paylocLogin, string paylocPassword)
        {
            _host           = host;
            _paylocLogin    = paylocLogin;
            _paylocPassword = paylocPassword;
        }

        // ── Connection ─────────────────────────────────────────────────────────────────

        public void Connect()
        {
            _session = new Session();          // Adjust ctor to match your DLL
            _session.Open(_host);
            Wait(ActionDelayMs);
            LoginPayloc();
        }

        private void LoginPayloc()
        {
            // Navigate to 111W and authenticate via Payloc method.
            // Adjust field coordinates / key sequences to match your AS400 login screen.
            SendText(_paylocLogin);
            SendKey(Key.Tab);
            SendText(_paylocPassword);
            SendKey(Key.Enter);
            Wait(ActionDelayMs * 2);
        }

        // ── Public API used by OrsResearcher ──────────────────────────────────────────

        /// <summary>
        /// Opens the ORS inquiry for a given ID and leaves the terminal on the
        /// first comment page. Returns false if the ORS could not be opened.
        /// </summary>
        public bool OpenOrs(string orsId)
        {
            // F9  →  FN = 34  →  ORS# = orsId  →  Enter
            SendKey(Key.F9);
            Wait(ActionDelayMs);

            SetField(FN_ROW, FN_COL, "34");
            SetField(ORS_ROW, ORS_COL, orsId);
            SendKey(Key.Enter);
            Wait(ActionDelayMs);

            // Verify we landed on a valid screen (not an error message).
            // Adjust error-text to whatever 111W shows for a bad ORS ID.
            string row1 = GetText(1, 1, 80);
            if (row1.Contains("NOT FOUND") || row1.Contains("INVALID") || row1.Contains("ERROR"))
                return false;

            return true;
        }

        /// <summary>
        /// Searches forward through comment pages (F8) for a line where
        /// off-id1 == "224" AND off-id2 == "007".
        /// Returns the date string from that line, or null if not found before
        /// the LAST PAGE sentinel.
        /// </summary>
        public string FindSentDate()
        {
            return ScanPages(scanForward: true, targetId1: "224", targetId2: "007");
        }

        /// <summary>
        /// Presses F7 and then scans forward for off-id1 == "007" AND off-id2 == "224"
        /// (reverse pair) to capture received date.
        /// </summary>
        public string FindReceivedDate()
        {
            // F7 goes backward one page – per spec we press it once then scan forward.
            SendKey(Key.F7);
            Wait(PageScrollDelayMs);

            return ScanPages(scanForward: true, targetId1: "007", targetId2: "224");
        }

        /// <summary>
        /// Advances to the next ORS by pressing F2, then waits for the main menu.
        /// </summary>
        public void GoToMainMenu()
        {
            // Scroll to last page first (F8 until sentinel) so F2 brings us home.
            ScrollToLastPage();

            SendKey(Key.F2);
            Wait(ActionDelayMs);

            // Wait until main menu is visible (up to ~5 s)
            int tries = 0;
            while (tries++ < 10)
            {
                if (GetText(1, 1, 80).Contains(MAIN_MENU_TEXT))
                    break;
                Wait(ActionDelayMs / 2);
            }
        }

        // ── Private helpers ───────────────────────────────────────────────────────────

        private string ScanPages(bool scanForward, string targetId1, string targetId2)
        {
            const int MAX_PAGES = 500; // safety cap
            int page = 0;

            while (page++ < MAX_PAGES)
            {
                // Check every row of the visible screen
                string result = ScanCurrentScreenForMatch(targetId1, targetId2);
                if (result != null)
                    return result;

                // Check for last-page sentinel on row 1
                if (GetText(1, 1, 80).Contains(LAST_PAGE_TEXT))
                    return null;

                // Scroll
                SendKey(scanForward ? Key.F8 : Key.F7);
                Wait(PageScrollDelayMs);
            }

            return null; // not found within page limit
        }

        private string ScanCurrentScreenForMatch(string targetId1, string targetId2)
        {
            // AS400 screens are typically 24 rows × 80 cols.
            for (int row = 1; row <= 24; row++)
            {
                // Read relevant columns from this row
                string field1 = GetText(row, OFFID1_COL, 3).Trim();
                string field2 = GetText(row, OFFID2_COL, 3).Trim();

                if (field1 == targetId1 && field2 == targetId2)
                {
                    // Grab the date from column 2 on the same row.
                    // AS400 dates are typically 8 chars: MM/DD/YY or MM/DD/YYYY
                    string dateStr = GetText(row, DATE_COL, 10).Trim();
                    return dateStr;
                }
            }
            return null;
        }

        private void ScrollToLastPage()
        {
            const int MAX_PAGES = 500;
            int page = 0;
            while (page++ < MAX_PAGES)
            {
                if (GetText(1, 1, 80).Contains(LAST_PAGE_TEXT))
                    break;
                SendKey(Key.F8);
                Wait(PageScrollDelayMs);
            }
        }

        // ── Low-level terminal wrappers ───────────────────────────────────────────────
        // Adjust these to match the exact method names in your UnetHelper.dll.

        private void SetField(int row, int col, string value)
        {
            _session.SetCursorPosition(row, col);
            _session.SendKeys(value);
            Wait(100);
        }

        private void SendText(string text)
        {
            _session.SendKeys(text);
            Wait(100);
        }

        private void SendKey(Key key)
        {
            _session.SendKey(key);
            Wait(ActionDelayMs);
        }

        /// <summary>
        /// Reads <paramref name="length"/> characters from the screen at (row, col).
        /// </summary>
        private string GetText(int row, int col, int length)
        {
            return _session.GetText(row, col, length) ?? string.Empty;
        }

        private void Wait(int ms)
        {
            Thread.Sleep(ms);
        }

        public void Dispose()
        {
            try { _session?.Close(); } catch { /* best-effort */ }
        }
    }

    // ── Placeholder key enum ──────────────────────────────────────────────────────────
    // Replace with the actual enum / constants from UnetHelper.dll.
    public enum Key
    {
        Enter, Tab, F2, F7, F8, F9
    }
}
