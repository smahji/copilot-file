// New file: src/MhiSession.cs

using System;
using System.Collections.Generic;
using System.Threading;

namespace ORSResearchTool
{
    /// <summary>
    /// Manages Session B — a separate UnetHelper session used exclusively for
    /// MHI screen lookups.  Opened once with SeamlessLogin (UnetID + UnetPass),
    /// stays alive for the entire run.
    /// </summary>
    public class MhiSession : IDisposable
    {
        private readonly string _sessionName = "B";
        private readonly string _unetId;
        private readonly string _unetPass;

        private bool _isConnected = false;

        // MHI screen sentinel
        private const string MHI_SCREEN_MARKER = "MH";   // autECLPSObj.GetText(2,2,2)
        private const int     SCREEN_ROWS       = 24;
        private const int     SCREEN_COLS       = 80;

        public MhiSession(string unetId, string unetPass)
        {
            _unetId   = unetId   ?? throw new ArgumentNullException(nameof(unetId));
            _unetPass = unetPass ?? throw new ArgumentNullException(nameof(unetPass));
        }

        // ── Connect (call once before processing loop) ────────────────────────

        /// <summary>
        /// Initialises emulator on Session B and performs seamless login.
        /// Mirrors: New UnetHelper.UnetHelper(), then SeamlessLogin(unetId, unetPass).
        /// </summary>
        public void Connect()
        {
            mdlUnet.InitializeEmulator(_sessionName);
            Thread.Sleep(200);
            mdlUnet.apiChk();

            // Seamless login on session B — UnetID + UnetPass only, no Payloc
            UnetHelper.UnetHelper unetObj = new UnetHelper.UnetHelper();
            unetObj.SeamlessLogin(_sessionName, _unetId, _unetPass);

            Thread.Sleep(500);
            mdlUnet.apiChk();
            _isConnected = true;
        }

        // ── MHI lookup ────────────────────────────────────────────────────────

        /// <summary>
        /// Opens the MHI screen for the given ICN, scrapes all remark codes and
        /// the adjuster ID from the last page.
        /// Returns false and leaves remarkCodes/adjusterId as null if MHI won't open.
        /// </summary>
        public bool TryReadMhi(
            string icn,
            out string remarkCodes,
            out string adjusterId)
        {
            remarkCodes = null;
            adjusterId  = null;

            if (!_isConnected)
                throw new InvalidOperationException("Call Connect() before TryReadMhi().");

            try
            {
                // ── Navigate to MHI ───────────────────────────────────────────
                mdlUnet.SendClear();
                Thread.Sleep(150);
                mdlUnet.SendClear();
                Thread.Sleep(150);

                // Control line: RET,I<icn>,M  at row 2, col 2
                // e.g. "RET,IFP58697495,M"
                mdlUnet.SendKeyes($"RET,I{icn},M", 2, 2);
                mdlUnet.SendEnter();
                Thread.Sleep(500);
                mdlUnet.apiChk();

                // Check MHI screen opened: (2,2,2) == "MH"
                if (autECLPSObj.GetText(2, 2, 2) != MHI_SCREEN_MARKER)
                    return false;   // MHI didn't open — caller will mark as NA and skip

                // ── Scroll through all pages, collect remark codes ────────────
                var codes        = new List<string>();
                string lastAdjId = null;
                const int MAX_PAGES = 200;

                for (int page = 0; page < MAX_PAGES; page++)
                {
                    lastAdjId = ScrapePageRemarkCodes(codes);

                    // Check for last page
                    string row24 = autECLPSObj.GetText(24, 1, 80) ?? string.Empty;
                    if (row24.Contains("NO MORE RECORDS") || row24.Contains("LAST PAGE"))
                        break;

                    // F8 to next page
                    autECLPSObj.SendKeys("[PF8]");
                    Thread.Sleep(300);
                    mdlUnet.apiChk();
                }

                remarkCodes = codes.Count > 0
                    ? string.Join(",", codes)
                    : null;

                adjusterId = lastAdjId;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ── Private helpers ───────────────────────────────────────────────────

        /// <summary>
        /// Reads every visible MHI row on the current page.
        /// MHI data rows are every even row (4,6,8…20) — mirrors ClaimProcessor.vb
        /// where YYYYY steps 4 To 20 Step 2.
        /// Remark codes are at (row, 64, 8); adjuster ID is read from the odd row
        /// below each entry.  Returns the last adjuster ID seen on this page.
        /// </summary>
        private string ScrapePageRemarkCodes(List<string> codes)
        {
            string lastAdjId = null;

            for (int row = 4; row <= 20; row += 2)
            {
                // Blank row means no more entries on this page
                string rowText = autECLPSObj.GetText(row, 1, 80) ?? string.Empty;
                if (rowText.Replace("-", "").Replace(" ", "").Length == 0)
                    break;

                // Remark codes at col 64, len 8 — clean up separators same as VB
                string rawCode = autECLPSObj.GetText(row, 64, 8) ?? string.Empty;
                string code    = rawCode
                    .Replace(" ", "")
                    .Replace("--", "")
                    .Replace(",,", ",")
                    .Trim(',');

                if (!string.IsNullOrWhiteSpace(code) && !codes.Contains(code))
                    codes.Add(code);

                // Adjuster ID is on the detail row (row+1) — read and keep last seen
                string adjRow = autECLPSObj.GetText(row + 1, 1, 80) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(adjRow))
                {
                    // Adjuster ID is typically the last token on the detail line.
                    // Adjust column/length if your MHI screen layout differs.
                    string adj = autECLPSObj.GetText(row + 1, 64, 10).Trim();
                    if (!string.IsNullOrWhiteSpace(adj))
                        lastAdjId = adj;
                }
            }

            return lastAdjId;
        }

        public void Dispose()
        {
            try
            {
                if (_isConnected)
                {
                    mdlUnet.SendClear();
                    mdlUnet.SendClear();
                }
            }
            catch { /* best-effort */ }
            finally { _isConnected = false; }
        }
    }
}