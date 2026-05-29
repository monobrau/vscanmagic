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
        Assert.True(layout.UsesMiscSubfolder);
        Assert.EndsWith(Path.Combine("2026 - Q1"), layout.OutputDirectory);
        Assert.EndsWith(Path.Combine("2026 - Q1", "Misc"), layout.TextOutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
        Assert.True(Directory.Exists(layout.TextOutputDirectory));
    }

    [Fact]
    public void ResolveVulnerabilityScansSubpath_AppendsNetworkDocumentationWhenMissing()
    {
        var result = ReportPathResolver.ResolveVulnerabilityScansSubpath("Contoso");
        Assert.Equal(
            Path.Combine("Contoso", "Network Documentation", "Vulnerability Scans"),
            result.Replace('/', Path.DirectorySeparatorChar));
    }

    [Fact]
    public void ResolveVulnerabilityScansSubpath_LeavesExistingPathUntouched()
    {
        var existing = Path.Combine("Client", "Network Documentation", "Vulnerability Scans");
        Assert.Equal(existing, ReportPathResolver.ResolveVulnerabilityScansSubpath(existing));
    }

    [Fact]
    public void ResolveQuarterFolderName_AlwaysUsesYearQuarterFromScanDate()
    {
        var clientPath = Path.Combine(_root, "Client", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(Path.Combine(clientPath, "2026 - Q1"));

        Assert.Equal("2026 - Q1", ReportPathResolver.ResolveQuarterFolderName(clientPath, "2026-02-10"));
        Assert.Equal("2026 - Q1", ReportPathResolver.ResolveQuarterFolderName(clientPath, "2026-03-31"));
    }

    [Fact]
    public void ResolveQuarterFolderName_IgnoresLegacyDatedQuarterFolders()
    {
        var clientPath = Path.Combine(_root, "Client", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(Path.Combine(clientPath, "2026 - Q2"));
        Directory.CreateDirectory(Path.Combine(clientPath, "2026 - Q2 2026-05-25"));

        var folder = ReportPathResolver.ResolveQuarterFolderName(clientPath, "2026-05-25");

        Assert.Equal("2026 - Q2", folder);
    }

    [Fact]
    public void Resolve_WithBasePath_UsesBareQuarterWhenLegacyDatedFoldersExist()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var clientFolder = Path.Combine(basePath, "Fabrikam Industries", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(Path.Combine(clientFolder, "2026 - Q2"));
        Directory.CreateDirectory(Path.Combine(clientFolder, "2026 - Q2 2026-05-25"));

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 0, "Fabrikam Industries Inc", "2026-05-25", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.EndsWith(Path.Combine("2026 - Q2"), layout.OutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
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
    public void Resolve_WithBasePathConfigured_UsesConfiguredBaseNotDeepFallbackWhenBaseMissing()
    {
        var fallback = Path.Combine(_root, "flat", "Global", "2026 - Q2", "Nested Client", "2026 - Q2");
        Directory.CreateDirectory(fallback);
        var configuredBase = Path.Combine(_root, "ReportsBase");
        var settings = new UserSettings
        {
            ReportsBasePath = configuredBase,
            LastOutputDirectory = fallback
        };

        var layout = _resolver.Resolve(settings, 0, "Unknown Client", "2026-01-01", fallback);

        Assert.True(layout.UsesStructuredPaths);
        Assert.True(layout.UsesMiscSubfolder);
        Assert.StartsWith(configuredBase, layout.OutputDirectory);
        Assert.EndsWith(Path.Combine("Unknown Client", "2026 - Q1"), layout.OutputDirectory);
        Assert.DoesNotContain("flat", layout.OutputDirectory);
        Assert.True(Directory.Exists(layout.OutputDirectory));
    }

    [Fact]
    public void Resolve_WithBasePath_DoesNotNestWhenFallbackIsPriorQuarterOutput()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        Directory.CreateDirectory(basePath);
        _folderMap.SetFolder(12345, "Accurate Metal Products/Network Documentation/Vulnerability Scans");
        var clientFolder = Path.Combine(basePath, "Accurate Metal Products", "Network Documentation", "Vulnerability Scans");
        var priorOutput = Path.Combine(clientFolder, "2026 - Q2 2026-05-26_101012");
        Directory.CreateDirectory(priorOutput);

        var settings = new UserSettings
        {
            ReportsBasePath = basePath,
            LastOutputDirectory = priorOutput
        };

        var layout = _resolver.Resolve(settings, 12345, "Accurate Metal Products Inc", "2026-05-26", priorOutput);

        Assert.True(layout.UsesStructuredPaths);
        Assert.StartsWith(clientFolder, layout.OutputDirectory);
        Assert.DoesNotContain(Path.Combine("2026 - Q2", "Accurate Metal Products"), layout.OutputDirectory);
    }

    [Fact]
    public void SanitizeMappedFolderPath_StripsQuarterSegments()
    {
        var input = @"Accurate Metal Products\Network Documentation\Vulnerability Scans\2026 - Q2\Accurate Metal Products\2026 - Q2";
        var sanitized = ReportPathResolver.SanitizeMappedFolderPath(input);
        Assert.Equal(@"Accurate Metal Products\Network Documentation\Vulnerability Scans", sanitized);
    }

    [Fact]
    public void NormalizeConfiguredBasePath_StripsQuarterFolderAndFindsGlobalRoot()
    {
        var nested = Path.Combine(_root, "Global", "2026 - Q2", "Global", "2026 - Q2");
        var normalized = ReportPathResolver.NormalizeConfiguredBasePath(nested);
        Assert.Equal(Path.Combine(_root, "Global"), normalized);
    }

    [Fact]
    public void NormalizeConfiguredBasePath_LeavesGlobalRootUnchanged()
    {
        var root = Path.Combine(_root, "Global");
        Assert.Equal(root, ReportPathResolver.NormalizeConfiguredBasePath(root));
    }

    [Fact]
    public void Resolve_WithBasePath_UsesClientQuarterWhenNoFolderMatch()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        Directory.CreateDirectory(basePath);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 0, "Brand New Client LLC", "2026-05-23", _root);

        Assert.True(layout.UsesStructuredPaths);
        Assert.True(layout.UsesMiscSubfolder);
        Assert.EndsWith(Path.Combine("Brand New Client LLC", "2026 - Q2"), layout.OutputDirectory);
        Assert.EndsWith(Path.Combine("Brand New Client LLC", "2026 - Q2", "Misc"), layout.TextOutputDirectory);
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
        Assert.EndsWith(Path.Combine("2026 - Q2", "Misc"), layout.TextOutputDirectory);
    }

    [Fact]
    public void LayoutForExistingDirectory_ReusesMiscSubfolder()
    {
        var quarter = Path.Combine(_root, "Client", "2026 - Q2");
        var misc = Path.Combine(quarter, "Misc");
        Directory.CreateDirectory(misc);

        var layout = ReportPathResolver.LayoutForExistingDirectory(quarter, "Client");

        Assert.Equal(quarter, layout.OutputDirectory);
        Assert.Equal(misc, layout.TextOutputDirectory);
        Assert.True(layout.UsesMiscSubfolder);
    }

    [Fact]
    public void GetDownloadDirectory_RoutesPendingEpssToMiscWhenStructured()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var clientFolder = Path.Combine(basePath, "Fabrikam Industries", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(clientFolder);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 12345, "Fabrikam Industries Inc", "2026-03-15", _root);

        Assert.Equal(layout.TextOutputDirectory, ReportPathResolver.GetDownloadDirectory(layout, "pending-epss"));
        Assert.Equal(layout.OutputDirectory, ReportPathResolver.GetDownloadDirectory(layout, "all-vulnerabilities"));
    }

    [Fact]
    public void GetDefaultManualOutputDirectory_UsesExportsWhenNoLastOutput()
    {
        var settings = new UserSettings();
        var path = ReportPathResolver.GetDefaultManualOutputDirectory(settings);

        Assert.EndsWith(Path.Combine("VScanMagic", "Exports"), path);
    }

    [Fact]
    public void GetSupplementalExportDirectory_UsesMiscWhenStructured()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var clientFolder = Path.Combine(basePath, "Fabrikam Industries", "Network Documentation", "Vulnerability Scans");
        Directory.CreateDirectory(clientFolder);

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = _resolver.Resolve(settings, 12345, "Fabrikam Industries Inc", "2026-03-15", _root);

        Assert.Equal(layout.OutputDirectory, ReportPathResolver.GetTopNReportDirectory(layout));
        Assert.Equal(layout.TextOutputDirectory, ReportPathResolver.GetSupplementalExportDirectory(layout));
        Assert.EndsWith("Misc", ReportPathResolver.GetSupplementalExportDirectory(layout));
    }

    [Fact]
    public void InferQuarterDirectoryFromSourceFile_UsesParentWhenFileInMisc()
    {
        var quarter = Path.Combine(_root, "Client", "2026 - Q2");
        var misc = Path.Combine(quarter, ReportPathResolver.MiscSubfolderName);
        Directory.CreateDirectory(misc);
        var file = Path.Combine(misc, "Pending EPSS.xlsx");
        File.WriteAllText(file, "x");

        var inferred = SessionOutputLayoutResolver.InferQuarterDirectoryFromSourceFile(file);

        Assert.Equal(quarter, inferred);
    }

    [Fact]
    public void ResolveForSession_PrefersDownloadFolderOverScanDateResolve()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var downloadQuarter = Path.Combine(basePath, "Acme Corp", "2026 - Q2");
        Directory.CreateDirectory(Path.Combine(downloadQuarter, ReportPathResolver.MiscSubfolderName));
        var sourceFile = Path.Combine(downloadQuarter, "All Vulnerabilities.xlsx");
        File.WriteAllText(sourceFile, "x");

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = SessionOutputLayoutResolver.ResolveForSession(
            _resolver,
            settings,
            "Acme Corp",
            "2026-01-15",
            0,
            sessionOutputDirectory: null,
            sourceFilePath: sourceFile);

        Assert.Equal(downloadQuarter, layout.OutputDirectory);
    }

    [Fact]
    public void ResolveForSession_UsesSessionOutputDirectoryWhenSet()
    {
        var basePath = Path.Combine(_root, "ReportsBase");
        var pinned = Path.Combine(basePath, "Acme Corp", "2026 - Q2");
        Directory.CreateDirectory(Path.Combine(pinned, ReportPathResolver.MiscSubfolderName));

        var settings = new UserSettings { ReportsBasePath = basePath };
        var layout = SessionOutputLayoutResolver.ResolveForSession(
            _resolver,
            settings,
            "Acme Corp",
            "2026-01-15",
            0,
            sessionOutputDirectory: pinned,
            sourceFilePath: null);

        Assert.Equal(pinned, layout.OutputDirectory);
        Assert.EndsWith("Misc", layout.TextOutputDirectory);
    }
}
