using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;

namespace VScanMagic.Tests;

public sealed class ReportPathResolverTests : IDisposable
{
    private readonly string _root;
    private readonly CompanyFolderMapService _folderMap;
    private readonly ReportPathResolver _resolver;

    public ReportPathResolverTests()
    {
        _root = Path.Combine(Path.GetTempPath(), "VScanMagicTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        _folderMap = new CompanyFolderMapService();
        _resolver = new ReportPathResolver(_folderMap);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_root))
                Directory.Delete(_root, recursive: true);
        }
        catch
        {
            // Best-effort cleanup for temp test dirs.
        }
    }

    [Fact]
    public void Resolve_WithoutBasePath_UsesFallbackFlatDirectory()
    {
        var fallback = Path.Combine(_root, "fallback");
        Directory.CreateDirectory(fallback);

        var settings = new UserSettings { LastOutputDirectory = fallback };
        var layout = _resolver.Resolve(settings, 12345, "Fabrikam Industries Inc", "2026-03-15", fallback);

        Assert.Equal(fallback, layout.OutputDirectory);
        Assert.Equal(fallback, layout.TextOutputDirectory);
        Assert.False(layout.UsesStructuredPaths);
        Assert.False(layout.UsesMiscSubfolder);
    }

    [Fact]
    public void Resolve_WithBasePath_CreatesQuarterFolder()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var clientFolder = Path.Combine(basePath, "Fabrikam Industries", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(clientFolder);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 12345, "Fabrikam Industries Inc", "2026-03-15", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.False(layout.UsesMiscSubfolder);
        Assert.EndsWith(Path.Combine("2026 - Q1"), layout.OutputDirectory);
        Assert.Equal(layout.OutputDirectory, layout.TextOutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
    }

    [Fact]
    public void ResolveVulnerabilityScansSubpath_AppendsNetworkDocumentationWhenMissing()
    {
        var result = ReportPathResolver.ResolveVulnerabilityScansSubpath("Contoso");
        Assert.Equal(Path.Combine("Contoso", "Network Documentation", "Vulnerability Scans"), result);
    }

    [Fact]
    public void ResolveVulnerabilityScansSubpath_LeavesExistingPathUntouched()
    {
        var existing = Path.Combine("Client", "Network Documentation", "Vulnerability Scans");
        Assert.Equal(existing, ReportPathResolver.ResolveVulnerabilityScansSubpath(existing));
    }

    [Fact]
    public void ResolveQuarterFolderName_AddsDateWhenQuarterExists()
    {
        var clientPath = Path.Combine(_root, "Client", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(Path.Combine(clientPath, "2026 - Q1"));

        var folder = ReportPathResolver.ResolveQuarterFolderName(clientPath, "2026-02-10");
        Assert.Equal("2026 - Q1 2026-02-10", folder);
    }

    [Fact]
    public void GetReportsPathPartial_IncludesCompanyAndNetworkDocumentation()
    {
        var fullPath = Path.Combine(
            _root,
            "Fabrikam Industries",
            "Network Documentation",
            "Vulnerability Scans",
            "2026 - Q1");

        var partial = ReportPathResolver.GetReportsPathPartial(fullPath, "Fabrikam Industries");
        Assert.Equal(@"Fabrikam Industries\Network Documentation\Vulnerability Scans\2026 - Q1", partial);
    }

    [Fact]
    public void GetSafeReportOutputPath_TruncatesLongCompanyName()
    {
        var targetDir = Path.Combine(_root, "out");
        Directory.CreateDirectory(targetDir);
        var longName = new string('A', 200);

        var path = ReportPathResolver.GetSafeReportOutputPath(
            targetDir,
            longName,
            " Top Ten Vulnerabilities Report_2026-03-15_120000",
            "docx");

        Assert.True(path.Length <= 250);
        Assert.EndsWith(".docx", path);
    }

    [Fact]
    public void Resolve_WithBasePathConfigured_UsesClientQuarterOnFallbackWhenBaseMissing()
    {
        var fallback = Path.Combine(_root, "flat");
        Directory.CreateDirectory(fallback);
        var settings = new UserSettings
        {
            ReportsBasePath = Path.Combine(_root, "missing-base"),
            LastOutputDirectory = fallback
        };

        var layout = _resolver.Resolve(settings, 0, "Unknown Client", "2026-01-01", fallback);

        Assert.True(layout.UsesStructuredPaths);
        Assert.False(layout.UsesMiscSubfolder);
        Assert.EndsWith(Path.Combine("Unknown Client", "2026 - Q1"), layout.OutputDirectory);
        Assert.Equal(layout.OutputDirectory, layout.TextOutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
    }

    [Fact]
    public void Resolve_WithBasePath_UsesClientQuarterWhenNoFolderMatch()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        Directory.CreateDirectory(basePath);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 0, "Brand New Client LLC", "2026-05-23", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.False(layout.UsesMiscSubfolder);
        Assert.EndsWith(Path.Combine("Brand New Client LLC", "2026 - Q2"), layout.OutputDirectory);
        Assert.Equal(layout.OutputDirectory, layout.TextOutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
    }

    [Fact]
    public void SanitizeClientFolderName_RemovesInvalidPathCharacters()
    {
        var sanitized = ReportPathResolver.SanitizeClientFolderName("Client: Name/With*Chars?");
        Assert.DoesNotContain(':', sanitized);
        Assert.DoesNotContain('/', sanitized);
        Assert.DoesNotContain('*', sanitized);
        Assert.DoesNotContain('?', sanitized);
    }

    [Fact]
    public void Resolve_WithBasePath_MatchesClientFolderByNameWithoutCompanyId()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var clientFolder = Path.Combine(basePath, "Fabrikam Industries", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(clientFolder);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 0, "Fabrikam Industries Inc", "2026-03-15", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.EndsWith(Path.Combine("2026 - Q1"), layout.OutputDirectory);
    }

    [Fact]
    public void Resolve_WithBasePath_UsesMappedFolderWithWindowsSeparators()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        Directory.CreateDirectory(basePath);

        _folderMap.SetFolder(42, @"Contoso\Network Documentation\Vulnerability Scans");

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 42, "Contoso LLC", "2026-03-15", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.EndsWith(Path.Combine("2026 - Q1"), layout.OutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
    }

    [Fact]
    public void Resolve_WithBasePath_CreatesMappedClientPathWhenMissing()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        Directory.CreateDirectory(basePath);
        _folderMap.SetFolder(12345, "Fabrikam Industries/Network Documentation/Vulnerability Scans");

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 12345, "Fabrikam Industries Inc", "2026-05-23", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.Contains("Fabrikam Industries", layout.OutputDirectory);
        Assert.Contains("Network Documentation", layout.OutputDirectory);
        Assert.Contains("Vulnerability Scans", layout.OutputDirectory);
        Assert.EndsWith(Path.Combine("2026 - Q2"), layout.OutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
        Assert.Equal(layout.OutputDirectory, layout.TextOutputDirectory);
    }
}
