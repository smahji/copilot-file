using System;

namespace ORSResearchTool
{
    public class OrsRecord
    {
        public string OrsId { get; set; }
        public DateTime? SentDate { get; set; }
        public DateTime? ReceivedDate { get; set; }
        public string Difference { get; set; }  // days as string, or "NA"
        public string Status { get; set; }       // "Found", "NA", "Error"
    }
}
