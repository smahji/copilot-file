using System;
using System.Collections.Generic;

namespace ORSResearchTool
{
    /// <summary>
    /// Iterates over a list of ORS IDs, drives the AS400 session for each,
    /// and returns populated OrsRecord results.
    /// </summary>
    public class OrsResearcher
    {
        private readonly AS400Session _session;

        public event Action<string> OnStatusUpdate;   // fires with progress text
        public event Action<int, int> OnProgress;     // fires with (current, total)

        public OrsResearcher(AS400Session session)
        {
            _session = session;
        }

        public List<OrsRecord> Research(IList<string> orsIds)
        {
            var results = new List<OrsRecord>();
            int total = orsIds.Count;

            for (int i = 0; i < total; i++)
            {
                string orsId = orsIds[i].Trim();
                OnProgress?.Invoke(i + 1, total);
                OnStatusUpdate?.Invoke($"[{i + 1}/{total}] Processing ORS: {orsId}");

                var record = new OrsRecord { OrsId = orsId };

                try
                {
                    bool opened = _session.OpenOrs(orsId);

                    if (!opened)
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        results.Add(record);
                        OnStatusUpdate?.Invoke($"  → ORS {orsId} not found – skipped.");
                        NavigateToMainMenu();
                        continue;
                    }

                    // ── Find Sent Date (224 → 007) ────────────────────────────────
                    string sentRaw = _session.FindSentDate();
                    DateTime sentDate;

                    if (sentRaw == null || !TryParseAs400Date(sentRaw, out sentDate))
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        results.Add(record);
                        OnStatusUpdate?.Invoke($"  → Sent date (224/007) not found – skipped.");
                        NavigateToMainMenu();
                        continue;
                    }

                    // ── Find Received Date (007 → 224) ────────────────────────────
                    string recvRaw = _session.FindReceivedDate();
                    DateTime recvDate;

                    if (recvRaw == null || !TryParseAs400Date(recvRaw, out recvDate))
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        results.Add(record);
                        OnStatusUpdate?.Invoke($"  → Received date (007/224) not found – skipped.");
                        NavigateToMainMenu();
                        continue;
                    }

                    // ── Calculate TAT ─────────────────────────────────────────────
                    int diffDays = (int)(recvDate - sentDate).TotalDays;

                    record.SentDate      = sentDate;
                    record.ReceivedDate  = recvDate;
                    record.Difference    = diffDays.ToString();
                    record.Status        = "Found";

                    OnStatusUpdate?.Invoke(
                        $"  → Sent: {sentDate:MM/dd/yyyy}  Received: {recvDate:MM/dd/yyyy}  TAT: {diffDays} days");
                }
                catch (Exception ex)
                {
                    record.Status     = "Error";
                    record.Difference = "NA";
                    OnStatusUpdate?.Invoke($"  → ERROR on {orsId}: {ex.Message}");
                }

                results.Add(record);
                NavigateToMainMenu();
            }

            return results;
        }

        private void NavigateToMainMenu()
        {
            try { _session.GoToMainMenu(); }
            catch { /* best-effort – already on main menu or session issue */ }
        }

        // ── Date parsing ─────────────────────────────────────────────────────────────
        // AS400 commonly outputs MM/DD/YY or MM/DD/YYYY.  Add more formats as needed.
        private static readonly string[] DATE_FORMATS = new[]
        {
            "MM/dd/yy",
            "MM/dd/yyyy",
            "M/d/yy",
            "M/d/yyyy",
            "MMddyy",
            "MMddyyyy"
        };

        private static bool TryParseAs400Date(string raw, out DateTime result)
        {
            result = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            raw = raw.Trim();
            foreach (var fmt in DATE_FORMATS)
            {
                if (DateTime.TryParseExact(raw, fmt,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out result))
                    return true;
            }

            // Fallback: let .NET try
            return DateTime.TryParse(raw, out result);
        }
    }
}
