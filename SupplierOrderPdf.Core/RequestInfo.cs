using System;

namespace SupplierOrderPdf.Core;

public class RequestInfo
{
    public int OrderId { get; set; }
    public DateTime? CreatedUtc { get; set; }
    public int? CreatedByUserId { get; set; }
    public string CreatedByDisplayName { get; set; } = string.Empty;
    public string CreatedByEmail { get; set; } = string.Empty;
    public string CreatedByPhone { get; set; } = string.Empty;
    public string PdfPath { get; set; } = string.Empty;
    public DateTime? SentUtc { get; set; }
    public int? SentByUserId { get; set; }
    public string SentByDisplayName { get; set; } = string.Empty;
    public string LastEmailTo { get; set; } = string.Empty;

    public DateTime? CreatedLocal => CreatedUtc?.ToLocalTime();
    public DateTime? SentLocal => SentUtc?.ToLocalTime();
}
