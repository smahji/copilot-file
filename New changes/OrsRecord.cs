// In: OrsRecord.cs
// Replace: the existing class body

public class OrsRecord
{
    public string    OrsId        { get; set; }
    public string    Icn          { get; set; }   // Feature 1: ICN from (4,44,10)
    public DateTime? SentDate     { get; set; }
    public string    SentComment  { get; set; }   // Feature 2: blue comment line after 224/007
    public DateTime? ReceivedDate { get; set; }
    public string    RecvComment  { get; set; }   // Feature 2: blue comment line after 007/224
    public string    Difference   { get; set; }
    public string    RemarkCodes  { get; set; }   // Feature 3: comma-separated from MHI
    public string    AdjusterId   { get; set; }   // Feature 3: last page of MHI
    public string    Status       { get; set; }   // "Found", "NA", "Error"
}