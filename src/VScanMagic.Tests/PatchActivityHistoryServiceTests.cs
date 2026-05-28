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

    [Fact]
    public void Record_PreservesOtherCompanyEntries()
    {
        for (var i = 0; i < 105; i++)
        {
            _service.Record(new PatchActivityEntry(
                100,
                $"job-100-{i}",
                "Application Patch",
                "Submitted",
                $"Entry {i}",
                null,
                null,
                DateTimeOffset.UtcNow.AddMinutes(-i),
                null));
        }

        _service.Record(new PatchActivityEntry(
            200,
            "job-200",
            "OS Patch",
            "Submitted",
            "Windows update",
            null,
            null,
            DateTimeOffset.UtcNow,
            null));

        var company200 = _service.GetEntries(200);
        Assert.Single(company200);
        Assert.Equal("job-200", company200[0].JobId);

        var company100 = _service.GetEntries(100, limit: 200);
        Assert.Equal(100, company100.Count);
    }

    [Fact]
    public void GetByJobIdAndUpdateEntry_PersistsVerification()
    {
        _service.Record(new PatchActivityEntry(
            100,
            "job-verify",
            "Application Patch",
            "Submitted",
            "Chrome on HOST1",
            "HOST1",
            null,
            DateTimeOffset.UtcNow,
            "Patch request accepted.",
            [10],
            [42],
            "Google Chrome",
            "148.0.1"));

        var found = _service.GetByJobId(100, "job-verify");
        Assert.NotNull(found);
        Assert.Equal("Submitted", found!.Status);

        var updated = found with
        {
            VersionCheckStatus = "Verified",
            Status = "Verified",
            VerificationSummary = "Verified: 1/1 at target.",
            VerifiedAt = DateTimeOffset.UtcNow
        };
        Assert.True(_service.UpdateEntry(updated));

        var reloaded = _service.GetByJobId(100, "job-verify");
        Assert.NotNull(reloaded);
        Assert.Equal("Verified", reloaded!.VersionCheckStatus);
        Assert.Equal("Verified: 1/1 at target.", reloaded.VerificationSummary);
    }
}
