using VScanMagic.ConnectSecure;
using VScanMagic.Core.Services;

namespace VScanMagic.Tests;

public sealed class PatchJobDisplayHelperTests
{
    [Fact]
    public void ResolveApplication_UsesProductField()
    {
        var job = new PatchJobEntry(
            "1",
            "Application Patch",
            "Queued",
            "Wireshark on host",
            "host",
            null,
            null,
            Product: "Wireshark");

        Assert.Equal("Wireshark", PatchJobDisplayHelper.ResolveApplication(job));
    }

    [Fact]
    public void ResolveApplication_ExtractsFromDescription_WhenProductMissing()
    {
        var job = new PatchJobEntry(
            "1",
            "Application",
            "Initiated",
            "Mozilla Firefox — Success: 0, Failed: 0, Pending: 2",
            "Roswell.dorks.lan",
            null,
            null);

        Assert.Equal("Mozilla Firefox", PatchJobDisplayHelper.ResolveApplication(job));
    }

    [Fact]
    public void BuildMergedDescription_PrefersLocalHostContext_OverRemoteCountsOnly()
    {
        var entry = new PatchActivityEntry(
            1,
            "local",
            "Application Patch",
            "Submitted",
            "Npcap on Zillah3.dorks.lan",
            "Zillah3.dorks.lan",
            null,
            DateTimeOffset.Now,
            null,
            Product: "Npcap");

        var remote = new PatchJobCorrelationHelper.ParsedConnectSecureJob(
            "cs",
            "Application",
            "Success",
            "Npcap — Success: 1, Failed: 0, Pending: 0",
            "Zillah3.dorks.lan",
            null,
            null,
            DateTimeOffset.Now,
            ProductName: "Npcap");

        var merged = PatchJobDisplayHelper.BuildMergedDescription(entry, remote, "Npcap");
        Assert.Contains("Zillah3", merged, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ResolveDetail_StripsDuplicateProductPrefix()
    {
        var job = new PatchJobEntry(
            "1",
            "Application",
            "Queued",
            "Wireshark on DC-05.dorks.lan",
            "DC-05.dorks.lan",
            null,
            null,
            Product: "Wireshark");

        Assert.Equal("DC-05.dorks.lan", PatchJobDisplayHelper.ResolveDetail(job));
    }
}
