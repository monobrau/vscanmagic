using VScanMagic.Core.Services;

namespace VScanMagic.Tests;

public sealed class PatchActivityHistoryServiceTests : IDisposable
{
    private readonly string _configDir;
    private readonly PatchActivityHistoryService _service;

    public PatchActivityHistoryServiceTests()
    {
        _configDir = Path.Combine(Path.GetTempPath(), "VScanMagicTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_configDir);
        _service = new PatchActivityHistoryService(_configDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_configDir))
            Directory.Delete(_configDir, recursive: true);
    }

    [Fact]
    public void RecordAndGetEntries_FiltersByCompany()
    {
        _service.Record(new PatchActivityEntry(
            100,
            "job-a",
            "Application Patch",
            "Submitted",
            "7-Zip on HOST1",
            "HOST1",
            null,
            DateTimeOffset.UtcNow,
            "Patch request accepted."));

        _service.Record(new PatchActivityEntry(
            200,
            "job-b",
            "OS Patch",
            "Submitted",
            "Windows update on HOST2",
            "HOST2",
            null,
            DateTimeOffset.UtcNow.AddMinutes(-5),
            "Patch request accepted."));

        var company100 = _service.GetEntries(100);
        Assert.Single(company100);
        Assert.Equal("job-a", company100[0].JobId);

        var company200 = _service.GetEntries(200);
        Assert.Single(company200);
        Assert.Equal("job-b", company200[0].JobId);
    }
}
