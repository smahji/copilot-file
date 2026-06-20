// In OrsResearcher.cs
// Replace:
//   private readonly AS400Session _session;
//   public OrsResearcher(AS400Session session)
//   { _session = session ?? ...; }
// With:

private readonly AS400Session _session;
private readonly MhiSession   _mhi;

public OrsResearcher(AS400Session session, MhiSession mhi)
{
    _session = session ?? throw new ArgumentNullException(nameof(session));
    _mhi     = mhi     ?? throw new ArgumentNullException(nameof(mhi));
}










// In OrsResearcher.cs — replace the try{} block body inside the for loop

if (!_session.OpenOrs(orsId))
{
    record.Status     = "NA";
    record.Difference = "NA";
    OnStatusUpdate?.Invoke("  → Not found – skipped.");
    SafeGoToMainMenu();
    CommitRecord(outputPath, record, results);
    continue;
}

// ── Feature 1: ICN ───────────────────────────────────────────────────
record.Icn = _session.ReadIcn();
OnStatusUpdate?.Invoke($"  → ICN: {record.Icn ?? "not found"}");

// ── Sent date + comment (224 → 007) ─────────────────────────────────
string sentRaw = _session.FindSentDate(out string sentComment);
if (sentRaw == null || !TryParseAs400Date(sentRaw, out DateTime sentDate))
{
    record.Status     = "NA";
    record.Difference = "NA";
    OnStatusUpdate?.Invoke("  → Sent date (224/007) not found – skipped.");
    SafeGoToMainMenu();
    CommitRecord(outputPath, record, results);
    continue;
}
record.SentDate    = sentDate;
record.SentComment = sentComment;

// ── Received date + comment (007 → 224) ─────────────────────────────
string recvRaw = _session.FindReceivedDate(out string recvComment);
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
record.RecvComment  = recvComment;

int diff        = (int)(recvDate - sentDate).TotalDays;
record.Difference = diff.ToString();

// ── Feature 3: MHI lookup on Session B ──────────────────────────────
if (!string.IsNullOrWhiteSpace(record.Icn))
{
    bool mhiOk = _mhi.TryReadMhi(record.Icn,
                                  out string remarkCodes,
                                  out string adjusterId);
    if (mhiOk)
    {
        record.RemarkCodes = remarkCodes;
        record.AdjusterId  = adjusterId;
        OnStatusUpdate?.Invoke(
            $"  → MHI: codes={remarkCodes ?? "none"}  adj={adjusterId ?? "none"}");
    }
    else
    {
        OnStatusUpdate?.Invoke("  → MHI not opened – skipping MHI fields.");
    }
}

record.Status = "Found";
OnStatusUpdate?.Invoke(
    $"  → Sent: {sentDate:MM/dd/yyyy}  Recv: {recvDate:MM/dd/yyyy}  TAT: {diff}d");