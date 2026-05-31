namespace VScanMagic.Review.Models;

public enum BulkReviewJobStatus
{
    Queued,
    Running,
    Completed,
    Failed,
    Cancelled
}

public enum BulkReviewItemPhase
{
    Pending,
    Downloading,
    Ingesting,
    Completed,
    Failed,
    Skipped
}

public sealed class BulkReviewJobItem
{
    public string CompanyId { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public bool IsRmitPlus { get; set; }
    public BulkReviewItemPhase Phase { get; set; } = BulkReviewItemPhase.Pending;
    public string? StatusMessage { get; set; }
    public string? ErrorMessage { get; set; }
    public string? SessionId { get; set; }
    public string? OutputDirectory { get; set; }
}

public sealed class BulkReviewJob
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string ScanDate { get; set; } = "";
    public string Presenter { get; set; } = "";
    public bool DefaultIsRmitPlus { get; set; }
    public BulkReviewJobStatus Status { get; set; } = BulkReviewJobStatus.Queued;
    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.Now;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.Now;
    public List<BulkReviewJobItem> Items { get; set; } = [];
    public string? ErrorMessage { get; set; }

    public int CompletedCount => Items.Count(i => i.Phase == BulkReviewItemPhase.Completed);
    public int FailedCount => Items.Count(i => i.Phase == BulkReviewItemPhase.Failed);
    public int TotalCount => Items.Count;
}

public sealed class BulkReviewJobRequest
{
    public string ScanDate { get; set; } = "";
    public string Presenter { get; set; } = "";
    public bool DefaultIsRmitPlus { get; set; }
    public IReadOnlyList<BulkReviewCompanySelection> Companies { get; set; } = [];
}

public sealed record BulkReviewCompanySelection(string CompanyId, string CompanyName, bool IsRmitPlus);
