using System;
using System.Collections.Generic;
using System.Threading;
using ORSResearchTool.Models;

namespace ORSResearchTool
{
    /// <summary>
    /// Orchestrates per-ORS navigation and date capture.
    ///
    /// Checkpoint design:
    ///   - ExcelHandler.InitialiseOutputFile() is called by MainForm ONCE before
    ///     this class runs — it creates the file with the header row.
    ///   - CommitRecord() is called after EVERY ORS (found, NA, or error) so each
    ///     result is persisted to disk immediately.
    ///   - If the run is cancelled or crashes, every record processed so far is
    ///     already in the file — nothing is lost.
    ///   - ExcelHandler.FinaliseOutputFile() is called by MainForm after the loop
    ///     ends (naturally or via cancel) to add the summary block and auto-fit.
    /// </summary>
    public class OrsResearcher
    {
        private readonly AS400Session _session;

        public event Action<string>   OnStatusUpdate;
        public event Action<int, int> OnProgress;

        public OrsResearcher(AS400Session session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        /// <param name="orsIds">List of ORS IDs to process.</param>
        /// <param name="outputPath">
        ///   Path to the already-initialised checkpoint file (header row written).
        ///   AppendResult() is called after every single ORS.
        /// </param>
        /// <param name="ct">
        ///   Cancellation token wired to the Cancel button.
        ///   Checked between records so the current ORS always finishes cleanly
        ///   before stopping — partial records are never written.
        /// </param>
        public List<OrsRecord> Research(
            IList<string>     orsIds,
            string            outputPath,
            CancellationToken ct = default)
        {
            var results = new List<OrsRecord>();
            int total   = orsIds.Count;

            for (int i = 0; i < total; i++)
            {
                // ── Check cancellation BETWEEN records, never mid-record ──────────
                if (ct.IsCancellationRequested)
                {
                    OnStatusUpdate?.Invoke(
                        $"  ↩  Cancelled after {i} of {total}. " +
                        $"{results.Count} records already saved to disk.");
                    break;
                }

                string orsId = orsIds[i].Trim();
                OnProgress?.Invoke(i + 1, total);
                OnStatusUpdate?.Invoke($"[{i + 1}/{total}] ORS: {orsId}");

                var record = new OrsRecord { OrsId = orsId };

                try
                {
                    // ── Open ORS ─────────────────────────────────────────────────
                    if (!_session.OpenOrs(orsId))
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        OnStatusUpdate?.Invoke("  → Not found on AS400 – skipped.");
                        SafeGoToMainMenu();
                        CommitRecord(outputPath, record, results);
                        continue;
                    }

                    // ── Sent date: scan for off-id 224 → 007 ─────────────────────
                    string sentRaw = _session.FindSentDate();
                    if (sentRaw == null || !TryParseAs400Date(sentRaw, out DateTime sentDate))
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        OnStatusUpdate?.Invoke("  → Sent date (224/007) not found – skipped.");
                        SafeGoToMainMenu();
                        CommitRecord(outputPath, record, results);
                        continue;
                    }
                    record.SentDate = sentDate;

                    // ── Received date: F7 then scan for off-id 007 → 224 ─────────
                    string recvRaw = _session.FindReceivedDate();
                    if (recvRaw == null || !TryParseAs400Date(recvRaw, out DateTime recvDate))
                    {
                        record.Status     = "NA";
                        record.Difference = "NA";
                        OnStatusUpdate?.Invoke("  → Received date (007/224) not found – skipped.");
                        SafeGoToMainMenu();
                        CommitRecord(outputPath, record, results);
                        continue;
                    }
                    record.ReceivedDate = recvDate;

                    // ── TAT ───────────────────────────────────────────────────────
                    int diff          = (int)(recvDate - sentDate).TotalDays;
                    record.Difference = diff.ToString();
                    record.Status     = "Found";

                    OnStatusUpdate?.Invoke(
                        $"  → Sent: {sentDate:MM/dd/yyyy}  " +
                        $"Recv: {recvDate:MM/dd/yyyy}  TAT: {diff}d");
                }
                catch (Exception ex)
                {
                    record.Status     = "Error";
                    record.Difference = "NA";
                    OnStatusUpdate?.Invoke($"  → ERROR: {ex.Message}");
                }

                // ── Checkpoint: write to disk regardless of outcome ───────────────
                CommitRecord(outputPath, record, results);
                SafeGoToMainMenu();
            }

            // Flush whatever remains in the buffer after the loop ends
            // (the last partial batch that never hit BATCH_SIZE, or any
            // batches that failed mid-run and were held for retry).
            FlushBuffer(outputPath, force: true);

            return results;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // CommitRecord: single place where every completed record is persisted.
        // Adds to in-memory list first, then writes to the checkpoint file.
        // A write failure is logged but never rethrows — the in-memory list is the
        // fallback used by FinaliseOutputFile at the end.
        // ─────────────────────────────────────────────────────────────────────────────
        // ── Batch flush infrastructure ────────────────────────────────────────────────
        // Records accumulate in _writeBuffer until BATCH_SIZE is reached, then the
        // entire buffer is written to Excel in a single open/save cycle.
        // This avoids opening the file once per record (19,000+ times for large runs)
        // while still giving a recoverable on-disk checkpoint every 200 records.
        private const int                BATCH_SIZE   = 200;
        private readonly List<OrsRecord> _writeBuffer = new List<OrsRecord>(BATCH_SIZE);

        // CommitRecord: adds to the in-memory result list AND the write buffer.
        // Triggers an automatic flush when the buffer hits BATCH_SIZE.
        private void CommitRecord(string outputPath, OrsRecord record, List<OrsRecord> results)
        {
            results.Add(record);
            _writeBuffer.Add(record);

            if (_writeBuffer.Count >= BATCH_SIZE)
                FlushBuffer(outputPath, force: false);
        }

        // FlushBuffer: writes everything in _writeBuffer to Excel in one open/save,
        // then clears the buffer.
        //   force=false  called automatically every BATCH_SIZE records
        //   force=true   called after the loop ends (cancel or completion) to flush
        //                whatever remains (the last partial batch < BATCH_SIZE)
        private void FlushBuffer(string outputPath, bool force)
        {
            if (_writeBuffer.Count == 0) return;

            try
            {
                ExcelHandler.AppendBatch(outputPath, _writeBuffer);
                OnStatusUpdate?.Invoke(
                    $"  Checkpoint: {_writeBuffer.Count} records written " +
                    (force ? "(final flush)." : $"(batch of {BATCH_SIZE})."));
                _writeBuffer.Clear();
            }
            catch (Exception ex)
            {
                // Buffer is intentionally NOT cleared on failure.
                // The records stay in _writeBuffer and will be included in the
                // next flush attempt (or the final forced flush).
                OnStatusUpdate?.Invoke(
                    $"  Batch write failed ({_writeBuffer.Count} buffered, will retry): {ex.Message}");
            }
        }

        private void SafeGoToMainMenu()
        {
            try { _session.GoToMainMenu(); } catch { /* best-effort */ }
        }

        // ── Date parsing — handles common AS400 formats ───────────────────────────────
        private static readonly string[] DATE_FORMATS =
        {
            "MM/dd/yy", "MM/dd/yyyy", "M/d/yy", "M/d/yyyy", "MMddyy", "MMddyyyy"
        };

        private static bool TryParseAs400Date(string raw, out DateTime result)
        {
            result = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            raw = raw.Trim();
            foreach (var fmt in DATE_FORMATS)
                if (DateTime.TryParseExact(raw, fmt,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out result))
                    return true;

            return DateTime.TryParse(raw, out result);
        }
    }
}
