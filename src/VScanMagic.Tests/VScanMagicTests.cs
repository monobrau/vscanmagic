using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.DependencyInjection;
using System.IO.Compression;
using System.Text.Json;
using Microsoft.AspNetCore.Components.Web;
using VScanMagic.ConnectSecure;
using VScanMagic.Core.Configuration;
using VScanMagic.Core.Models;
using VScanMagic.Data;
using VScanMagic.Data.Parsing;
using VScanMagic.Data.Scoring;
using VScanMagic.Core.Nvd;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Services;
using VScanMagic.Review.Models;
using VScanMagic.Review.Storage;
using VScanMagic.Reports;
using VScanMagic.Web.Services;

namespace VScanMagic.Tests;

public class RiskScoreCalculatorTests
{
    [Fact]
    public void CompositeRiskScore_UsesSeverityWeightsAndEpss()
    {
        var options = new VScanMagicOptions();
        var score = Core.Risk.RiskScoreCalculator.GetCompositeRiskScore(2, 1, 0, 0, 0.5, "Test App", 3, options);
        var expected = Math.Round((2 * 0.9 + 1 * 0.8) * 1.5, 2);
        Assert.Equal(expected, score);
    }

    [Fact]
    public void CompositeRiskScore_BoostsCveOnlyFindingsBySeverityBand()
    {
        var options = new VScanMagicOptions();
        var cveScore = RiskScoreCalculator.GetCompositeRiskScore(1, 0, 0, 0, 0.85, "CVE-2021-33558", 1, options);
        var appScore = RiskScoreCalculator.GetCompositeRiskScore(0, 20, 0, 0, 0.3, "Google Chrome", 20, options);

        Assert.True(cveScore > appScore);
        Assert.Equal(Math.Round(9.0 * 1.5 * 1.85, 2), cveScore);
    }
}

public class TopVulnerabilityScorerTests
{
    [Fact]
    public void GetTopVulnerabilities_ReturnsRankedList()
    {
        var options = new VScanMagicOptions();
        var scorer = new TopVulnerabilityScorer(options);
        var data = new List<VulnerabilityRecord>
        {
            new() { Product = "Adobe Reader", HostName = "PC1", Critical = 2, High = 1, VulnerabilityCount = 3, EpssScore = 0.8 },
            new() { Product = "Chrome", HostName = "PC1", High = 5, VulnerabilityCount = 5, EpssScore = 0.2 },
            new() { Product = "Firefox", HostName = "PC2", Medium = 3, VulnerabilityCount = 3, EpssScore = 0.1 }
        };

        var top = scorer.GetTopVulnerabilities(data, new ReportFilters { TopN = 2, IncludeCritical = true, IncludeHigh = true, IncludeMedium = true });
        Assert.Equal(2, top.Count);
        Assert.Contains(top, t => t.Product == "Adobe Reader");
        Assert.True(top[0].RiskScore >= top[1].RiskScore);
    }

    [Fact]
    public void GetTopVulnerabilities_IncludesCveOnlyProducts()
    {
        var options = new VScanMagicOptions();
        var scorer = new TopVulnerabilityScorer(options);
        var data = new List<VulnerabilityRecord>
        {
            new() { Product = "CVE-2021-33558", HostName = "PC1", Critical = 1, VulnerabilityCount = 1, EpssScore = 0.9, Cve = "CVE-2021-33558" },
            new() { Product = "Google Chrome", HostName = "PC1", High = 3, VulnerabilityCount = 3, EpssScore = 0.5, Cve = "CVE-2024-1234" }
        };

        var top = scorer.GetTopVulnerabilities(data, new ReportFilters { TopN = 10, IncludeCritical = true, IncludeHigh = true });
        Assert.Equal(2, top.Count);
        Assert.Contains(top, item => item.Product == "CVE-2021-33558" && item.CveIds == "CVE-2021-33558");
        Assert.Contains(top, item => item.Product == "Google Chrome" && item.CveIds == "CVE-2024-1234");
    }

    [Fact]
    public void ScoreForReview_LoadsAllFiltered_AndAutoExportsApplicationOnly()
    {
        var options = new VScanMagicOptions();
        var scorer = new TopVulnerabilityScorer(options);
        var data = new List<VulnerabilityRecord>
        {
            new() { Product = "Google Chrome", HostName = "PC1", High = 20, VulnerabilityCount = 20, EpssScore = 0.5 },
            new() { Product = "Mozilla Firefox", HostName = "PC1", High = 15, VulnerabilityCount = 15, EpssScore = 0.4 },
            new() { Product = "SMB Signing", HostName = "(Network scan)", Source = "Network", High = 1, VulnerabilityCount = 1 },
            new() { Product = "CVE-2021-33558", HostName = "PC2", Source = "Registry", Critical = 1, VulnerabilityCount = 1, EpssScore = 0.8 }
        };

        var scored = scorer.ScoreForReview(data, new ReportFilters { TopN = 1, IncludeHigh = true, IncludeCritical = true });
        Assert.Equal(4, scored.AllFiltered.Count);
        Assert.Single(scored.AutoExportApplication);
        Assert.Equal("Google Chrome", scored.AutoExportApplication[0].Product);
    }

    [Fact]
    public void GetTopVulnerabilities_PrioritizesCriticalCveOverBulkHighFindings()
    {
        var options = new VScanMagicOptions { MinHighSeverityCveInTopN = 0 };
        var scorer = new TopVulnerabilityScorer(options);
        var data = new List<VulnerabilityRecord>
        {
            new() { Product = "CVE-2021-33558", HostName = "PC1", Critical = 1, VulnerabilityCount = 1, EpssScore = 0.85, Cve = "CVE-2021-33558" },
            new() { Product = "Google Chrome", HostName = "PC1", High = 20, VulnerabilityCount = 20, EpssScore = 0.3 },
            new() { Product = "Mozilla Firefox", HostName = "PC2", High = 15, VulnerabilityCount = 15, EpssScore = 0.25 }
        };

        var top = scorer.GetTopVulnerabilities(data, new ReportFilters { TopN = 2, IncludeCritical = true, IncludeHigh = true });

        Assert.Equal(2, top.Count);
        Assert.Equal("CVE-2021-33558", top[0].Product);
        Assert.Equal("Google Chrome", top[1].Product);
    }

    [Fact]
    public void CreateFromScoredResult_HonorsExportTopN_WhenScoredWithUnlimitedPool()
    {
        var factory = new ReviewSessionFactory(new RemediationRuleService(), new VScanMagicOptions());
        var applications = Enumerable.Range(1, 15)
            .Select(i => new TopVulnerability
            {
                Source = "Application",
                Product = $"App{i}",
                High = 16 - i,
                RiskScore = 16 - i,
                VulnCount = 1
            })
            .ToList();

        var scored = new ScoredVulnerabilityResult
        {
            AllFiltered = applications,
            AutoExportApplication = applications
        };

        var session = factory.CreateFromScoredResult("Acme", "2026-05-23", scored, "Tech", exportTopN: 10);

        Assert.Equal(10, ReviewSessionRanker.GetExportFindings(session).Count);
        Assert.Equal(15, session.Findings.Count);
    }

    [Fact]
    public void CreateFromScoredResult_KeepsNetworkAndRegistryAsCandidates()
    {
        var factory = new ReviewSessionFactory(new RemediationRuleService(), new VScanMagicOptions());
        var scored = new ScoredVulnerabilityResult
        {
            AllFiltered =
            [
                new TopVulnerability { Source = "Application", Product = "Google Chrome", High = 5, RiskScore = 10, EpssScore = 0.5, VulnCount = 5 },
                new TopVulnerability { Source = "Network", Product = "SMB Signing", High = 1, RiskScore = 2, VulnCount = 1 },
                new TopVulnerability { Source = "Registry", Product = "CVE-2021-33558", Critical = 1, RiskScore = 8, EpssScore = 0.7, VulnCount = 1, CveIds = "CVE-2021-33558" }
            ],
            AutoExportApplication =
            [
                new TopVulnerability { Source = "Application", Product = "Google Chrome", High = 5, RiskScore = 10, EpssScore = 0.5, VulnCount = 5 }
            ]
        };

        var session = factory.CreateFromScoredResult("Acme", "2026-05-23", scored, "Tech", exportTopN: 10);
        Assert.Single(ReviewSessionRanker.GetExportFindings(session));
        Assert.Equal(2, ReviewCandidatePool.CountCandidates(session));
        Assert.Equal(1, ReviewCandidatePool.CountCandidates(session, VulnerabilitySourceHelper.Network));
        Assert.Equal(1, ReviewCandidatePool.CountCandidates(session, VulnerabilitySourceHelper.Registry));
    }

    [Fact]
    public void PromoteToExport_AllowsManualNetworkItemWithoutRemovingOthers()
    {
        var session = new ReviewSession { ExportTopN = 1 };
        session.Findings.Add(new ReviewFinding { OriginalRank = 1, Product = "Chrome", Source = "Application", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 2, Product = "SMB", Source = "Network", IncludeInExport = false });

        ReviewSessionRanker.PromoteToExport(session, session.Findings[1]);

        Assert.Equal(2, ReviewSessionRanker.GetExportFindings(session).Count);
        Assert.True(session.Findings[1].ManuallyPromoted);
    }

    [Fact]
    public void CreateFromScoredResult_AutoUpdateBrowser_PrefersRuleGuidanceOverVersionFix()
    {
        var factory = new ReviewSessionFactory(new RemediationRuleService(), new VScanMagicOptions());
        var scored = new ScoredVulnerabilityResult
        {
            AllFiltered =
            [
                new TopVulnerability
                {
                    Source = "Application",
                    Product = "Brave",
                    Fix = "1.67.123",
                    High = 5,
                    RiskScore = 10,
                    EpssScore = 0.5,
                    VulnCount = 5
                }
            ],
            AutoExportApplication =
            [
                new TopVulnerability
                {
                    Source = "Application",
                    Product = "Brave",
                    Fix = "1.67.123",
                    High = 5,
                    RiskScore = 10,
                    EpssScore = 0.5,
                    VulnCount = 5
                }
            ]
        };

        var session = factory.CreateFromScoredResult("Acme", "2026-05-23", scored, "Tech", exportTopN: 10);
        var finding = Assert.Single(session.Findings);

        Assert.DoesNotContain("1.67.123", finding.RevisedRemediation);
        Assert.NotEqual("Update to version 1.67.123 or later", finding.RevisedRemediation);
        Assert.Contains("1.67.123", finding.OriginalFix);
    }
}

public class ExcelReaderTests
{
    [Fact]
    public void ReadFromFile_ParsesAllVulnerabilitiesFormat()
    {
        var path = CreateSampleWorkbook();
        try
        {
            var reader = new ExcelVulnerabilityReader();
            var records = reader.ReadFromFile(path);
            Assert.NotEmpty(records);
            Assert.Contains(records, r => r.Product == "TestProduct");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadFromFile_ParsesJsonArrayProductNames()
    {
        var path = CreateJsonProductWorkbook();
        try
        {
            var reader = new ExcelVulnerabilityReader();
            var records = reader.ReadFromFile(path);
            Assert.Contains(records, r => r.Product == "Microsoft Edge");
            Assert.Contains(records, r => r.Product == "Google Chrome");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadFromFile_RejectsTopTenDataExport()
    {
        var path = CreateTopTenExportWorkbook();
        try
        {
            var reader = new ExcelVulnerabilityReader();
            var ex = Assert.Throws<InvalidOperationException>(() => reader.ReadFromFile(path));
            Assert.Contains("Top Ten Data export", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string CreateTopTenExportWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"vscanmagic_test_{Guid.NewGuid():N}.xlsx");
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Top Findings");
        ws.Cell(1, 1).Value = "Rank";
        ws.Cell(1, 2).Value = "Product";
        ws.Cell(2, 1).Value = 1;
        ws.Cell(2, 2).Value = "Adobe Reader";
        wb.SaveAs(path);
        return path;
    }

    private static string CreateJsonProductWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"vscanmagic_test_{Guid.NewGuid():N}.xlsx");
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("All Vulnerabilities");
        ws.Cell(1, 1).Value = "Product Name";
        ws.Cell(1, 2).Value = "Host Name";
        ws.Cell(1, 3).Value = "IP Address";
        ws.Cell(1, 4).Value = "Severity";
        ws.Cell(1, 5).Value = "EPSS Score";
        ws.Cell(2, 1).Value = "[\"Microsoft Edge\", \"Google Chrome\"]";
        ws.Cell(2, 2).Value = "HOST01";
        ws.Cell(2, 3).Value = "10.0.0.1";
        ws.Cell(2, 4).Value = "Critical";
        ws.Cell(2, 5).Value = 0.75;
        wb.SaveAs(path);
        return path;
    }

    private static string CreateSampleWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"vscanmagic_test_{Guid.NewGuid():N}.xlsx");
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("All Vulnerabilities");
        ws.Cell(1, 1).Value = "Product Name";
        ws.Cell(1, 2).Value = "Host Name";
        ws.Cell(1, 3).Value = "IP Address";
        ws.Cell(1, 4).Value = "Severity";
        ws.Cell(1, 5).Value = "EPSS Score";
        ws.Cell(1, 6).Value = "Solution";
        ws.Cell(2, 1).Value = "TestProduct";
        ws.Cell(2, 2).Value = "HOST01";
        ws.Cell(2, 3).Value = "10.0.0.1";
        ws.Cell(2, 4).Value = "Critical";
        ws.Cell(2, 5).Value = 0.75;
        ws.Cell(2, 6).Value = "Upgrade to latest version";
        wb.SaveAs(path);
        return path;
    }
}

public class HostnameUsernameMatcherTests
{
    [Theory]
    [InlineData("AMI-W11-7.corp.local", "AMI-W11-7")]
    [InlineData("WORKSTATION", "WORKSTATION")]
    public void TryResolveUser_MatchesFqdnAndShortName(string storedHost, string lookupHost)
    {
        var index = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        HostnameUsernameMatcher.RegisterHost(index, storedHost, "jsmith");

        Assert.True(HostnameUsernameMatcher.TryResolveUser(index, lookupHost, out var user));
        Assert.Equal("jsmith", user);
    }

    [Fact]
    public void ApplyEmptyUsernames_OnlyFillsBlankUsernames()
    {
        var session = new Review.Models.ReviewSession
        {
            Findings =
            [
                new Review.Models.ReviewFinding
                {
                    AffectedSystems =
                    [
                        new Review.Models.ReviewAffectedSystem { HostName = "PC1", Username = "existing" },
                        new Review.Models.ReviewAffectedSystem { HostName = "PC2" }
                    ]
                }
            ]
        };

        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["PC1"] = "ignored",
            ["PC2"] = "newuser"
        };

        var updated = Review.ReviewUsernameLookup.ApplyEmptyUsernames(session, lookup);

        Assert.Equal(1, updated);
        Assert.Equal("existing", session.Findings[0].AffectedSystems[0].Username);
        Assert.Equal("newuser", session.Findings[0].AffectedSystems[1].Username);
    }
}

public class ReviewSessionRepositoryTests
{
    [Fact]
    public async Task SaveAndGet_RoundTripsSession()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vscanmagic_test_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        Environment.SetEnvironmentVariable("XDG_CONFIG_HOME", dir);

        using var repo = new Review.Storage.SqliteReviewSessionRepository(dir);
        var session = new Review.Models.ReviewSession
        {
            ClientName = "Test Co",
            ScanDate = "2026-01-01",
            Findings = [new Review.Models.ReviewFinding { Rank = 1, Product = "App", RevisedRemediation = "Fix it" }]
        };
        await repo.SaveAsync(session);
        var loaded = await repo.GetAsync(session.Id);
        Assert.NotNull(loaded);
        Assert.Equal("Test Co", loaded!.ClientName);
        Assert.Single(loaded.Findings);
    }

    [Fact]
    public async Task DeleteAsync_RemovesSession()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vscanmagic_test_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);

        using var repo = new Review.Storage.SqliteReviewSessionRepository(dir);
        var session = new Review.Models.ReviewSession { ClientName = "Delete Me", ScanDate = "2026-01-01" };
        await repo.SaveAsync(session);
        await repo.DeleteAsync(session.Id);

        Assert.Null(await repo.GetAsync(session.Id));
    }

    [Fact]
    public async Task ListAsync_ExcludesArchivedUnlessRequested()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vscanmagic_test_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);

        using var repo = new Review.Storage.SqliteReviewSessionRepository(dir);
        var active = new Review.Models.ReviewSession { ClientName = "Active", ScanDate = "2026-01-01" };
        var archived = new Review.Models.ReviewSession
        {
            ClientName = "Archived",
            ScanDate = "2026-01-01",
            ArchivedAt = DateTimeOffset.UtcNow
        };
        await repo.SaveAsync(active);
        await repo.SaveAsync(archived);

        var visible = await repo.ListAsync(includeArchived: false);
        var all = await repo.ListAsync(includeArchived: true);

        Assert.Single(visible);
        Assert.Equal("Active", visible[0].ClientName);
        Assert.Equal(2, all.Count);
    }
}

public class ReviewExportLabelsTests
{
    [Fact]
    public void GetReportTitle_UsesActualExportCountWhenBelowTopN()
    {
        var session = new ReviewSession
        {
            ExportTopN = 10,
            Findings =
            [
                new ReviewFinding { IncludeInExport = true, OriginalRank = 1, Product = "A" },
                new ReviewFinding { IncludeInExport = true, OriginalRank = 2, Product = "B" }
            ]
        };

        Assert.Equal("Top 2 Vulnerabilities Report", ReviewExportLabels.GetReportTitle(session));
        Assert.Equal("Top 2", ReviewExportLabels.GetTopNLabel(session));
    }

    [Fact]
    public void GetReportTitle_UsesTopTenWhenExportCountIsTen()
    {
        var session = new ReviewSession
        {
            ExportTopN = 10,
            Findings = Enumerable.Range(1, 10)
                .Select(i => new ReviewFinding { IncludeInExport = true, OriginalRank = i, Product = $"P{i}" })
                .ToList()
        };

        Assert.Equal("Top Ten Vulnerabilities Report", ReviewExportLabels.GetReportTitle(session));
        Assert.Equal("Top Ten", ReviewExportLabels.GetTopNLabel(session));
    }
}

public class DocxExporterTests
{
    [Fact]
    public void Export_CreatesDocxFile()
    {
        var session = new Review.Models.ReviewSession
        {
            ClientName = "Acme",
            ScanDate = "2026-05-01",
            Presenter = "Tech",
            Findings =
            [
                new Review.Models.ReviewFinding
                {
                    Rank = 1, Product = "App", RiskScore = 5.5, Epss = 0.3,
                    OriginalRemediation = "Update", RevisedRemediation = "Update next week"
                }
            ]
        };
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        try
        {
            new Reports.DocxReviewExporter(new RemediationRuleService()).Export(session, path);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 1000);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

public class ConnectSecureAuthTests
{
    [Fact]
    public void TryParseAuthResponse_ReadsNestedToken()
    {
        using var doc = JsonDocument.Parse("""{"data":{"access_token":"abc","user_id":"42"}}""");
        var ok = ConnectSecureClient.TryParseAuthResponse(doc.RootElement, out var token, out var userId, out var error);
        Assert.True(ok);
        Assert.Equal("abc", token);
        Assert.Equal("42", userId);
        Assert.Null(error);
    }

    [Fact]
    public void IsConnectSecureEncodedSecret_RecognizesPortalFormat()
    {
        Assert.True(ConnectSecureCredentialsHelper.IsConnectSecureEncodedSecret("Z0FBQUFBQnBCTlVLbUY0TEZxdWJYc20xcjUtdEtRU3BPWm1zdF9wbmZCZE1HSEhjYTkxclRmd1V0TmJBU3NKM212RkRaUXZpX1EyRVhFX1FNMXRXNC0yV0VqZkQwQUpqT1ZiQUhTWE5oaGI5cjN5WE0yWExmTXdqUGx0RHFCdzNNbVNSRTVQMFhfM2p1ekdkYUZWT0FzRkh4QWhKdjJGWVFrVEF1Z1JOVWpnUXl4NjFnUUhhZDlJPQ=="));
        Assert.False(ConnectSecureCredentialsHelper.IsConnectSecureEncodedSecret("short-secret"));
    }

    [Fact]
    public void TryParseAuthResponse_ReturnsMessageWhenUnauthorized()
    {
        using var doc = JsonDocument.Parse("""{"message":"Failed to authorize","status":false}""");
        var ok = ConnectSecureClient.TryParseAuthResponse(doc.RootElement, out _, out _, out var error);
        Assert.False(ok);
        Assert.Contains("Failed to authorize", error);
    }

    [Fact]
    public void TryParseAuthResponse_AcceptsTokenDespiteErrorMessage()
    {
        using var doc = JsonDocument.Parse("""{"message":"Failed to create customer","access_token":"abc","user_id":"9"}""");
        var ok = ConnectSecureClient.TryParseAuthResponse(doc.RootElement, out var token, out var userId, out _);
        Assert.True(ok);
        Assert.Equal("abc", token);
        Assert.Equal("9", userId);
    }
}

public class ReviewSessionRankerTests
{
    [Fact]
    public void Rebalance_PromotesNextFindingWhenOneIsExcluded()
    {
        var session = new ReviewSession { ExportTopN = 2 };
        session.Findings.Add(new ReviewFinding { OriginalRank = 1, Rank = 1, Product = "A", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 2, Rank = 2, Product = "B", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 3, Rank = 3, Product = "C", IncludeInExport = false });

        session.Findings[0].IncludeInExport = false;
        session.Findings[0].ExcludedFromExport = true;
        ReviewSessionRanker.Rebalance(session);

        Assert.Equal(["B", "C"], ReviewSessionRanker.GetExportFindings(session).Select(f => f.Product).ToArray());
        Assert.Equal(1, ReviewSessionRanker.GetExportRank(session, session.Findings[1]));
        Assert.Equal(2, ReviewSessionRanker.GetExportRank(session, session.Findings[2]));
    }

    [Fact]
    public void MarkConnectSecureSuppressed_RemovesFromExportAndSkipsRebalancePromotion()
    {
        var session = new ReviewSession { ExportTopN = 2 };
        session.Findings.Add(new ReviewFinding { OriginalRank = 1, Rank = 1, Product = "A", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 2, Rank = 2, Product = "B", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 3, Rank = 3, Product = "C", IncludeInExport = false });

        ReviewSessionRanker.MarkConnectSecureSuppressed(session, session.Findings[0], "False positive", "test");

        Assert.True(session.Findings[0].ConnectSecureSuppressed);
        Assert.Equal(FindingStatus.WontFix, session.Findings[0].Status);
        Assert.Equal(["B", "C"], ReviewSessionRanker.GetExportFindings(session).Select(f => f.Product).ToArray());
        Assert.True(session.Findings[0].ConnectSecureSuppressed);
        Assert.Equal(0, ReviewCandidatePool.CountCandidates(session));
    }

    [Fact]
    public void MarkConnectSecureUnsuppressed_ClearsSuppressStateAndRestoresOpenStatus()
    {
        var session = new ReviewSession { ExportTopN = 2 };
        var finding = new ReviewFinding
        {
            OriginalRank = 1,
            Rank = 1,
            Product = "A",
            ConnectSecureSuppressed = true,
            ConnectSecureSuppressRecordId = 99,
            ConnectSecureProblemId = 123,
            SuppressionReason = "False positive",
            SuppressionComments = "test",
            Status = FindingStatus.WontFix,
            ExcludedFromExport = true
        };
        session.Findings.Add(finding);

        ReviewSessionRanker.MarkConnectSecureUnsuppressed(session, finding);

        Assert.False(finding.ConnectSecureSuppressed);
        Assert.Null(finding.ConnectSecureSuppressRecordId);
        Assert.Null(finding.SuppressionReason);
        Assert.Equal(FindingStatus.Open, finding.Status);
        Assert.False(finding.ExcludedFromExport);
        Assert.Equal(123, finding.ConnectSecureProblemId);
    }

    [Fact]
    public void Rebalance_DoesNotPromoteSuppressedReserve()
    {
        var session = new ReviewSession { ExportTopN = 1 };
        session.Findings.Add(new ReviewFinding { OriginalRank = 1, Rank = 1, Product = "A", IncludeInExport = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 2, Rank = 2, Product = "B", IncludeInExport = false, ConnectSecureSuppressed = true });
        session.Findings.Add(new ReviewFinding { OriginalRank = 3, Rank = 3, Product = "C", IncludeInExport = false });

        session.Findings[0].IncludeInExport = false;
        session.Findings[0].ExcludedFromExport = true;
        ReviewSessionRanker.Rebalance(session);

        Assert.Equal(["C"], ReviewSessionRanker.GetExportFindings(session).Select(f => f.Product).ToArray());
    }
}

public class SuppressibleProblemMatcherTests
{
    [Fact]
    public void Match_FindsExactCveProblem()
    {
        var entries = new List<SuppressibleProblemEntry>
        {
            new(42, "CVE-2021-33558", 3),
            new(43, "CVE-2024-1234", 1)
        };

        var match = SuppressibleProblemMatcher.Match(entries, ["CVE-2021-33558"]);

        Assert.True(match.HasMatch);
        Assert.Equal(ReviewSuppressTargetKind.Problem, match.Entry!.Kind);
        Assert.Equal(42, match.Entry.Id);
    }
}

public class SuppressibleRemediationMatcherTests
{
    [Fact]
    public void Match_PrefersExactProductName()
    {
        var entries = new List<SuppressibleRemediationEntry>
        {
            new(10, "Google Chrome", "fix", "High", "", true, 5),
            new(11, "Chrome", "fix", "High", "", true, 2)
        };

        var match = SuppressibleRemediationMatcher.Match(entries, "Google Chrome");

        Assert.True(match.HasMatch);
        Assert.Equal(10, match.Entry!.SolutionId);
    }

    [Fact]
    public void Match_UsesMajorVersionWhenNamesDifferByPatchLevel()
    {
        var entries = new List<SuppressibleRemediationEntry>
        {
            new(10, "MongoDB 3.4", "fix", "High", "", true, 5)
        };

        var match = SuppressibleRemediationMatcher.Match(entries, "MongoDB 3.4.24");

        Assert.True(match.HasMatch);
        Assert.Equal(10, match.Entry!.SolutionId);
    }

    [Fact]
    public void Match_ReturnsAmbiguousWhenMultipleEqualScores()
    {
        var entries = new List<SuppressibleRemediationEntry>
        {
            new(10, "Widget", "fix", "High", "", true, 5),
            new(11, "Widget", "fix", "High", "", true, 3)
        };

        var match = SuppressibleRemediationMatcher.Match(entries, "Widget");

        Assert.False(match.HasMatch);
        Assert.Equal(2, match.AmbiguousMatches.Count);
    }
}

public class HostVulnerabilityReportExporterTests
{
    [Fact]
    public void Export_CreatesPdfAndXlsx()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vscanmagic_host_export_" + Guid.NewGuid().ToString("N"));
        var hosts = new List<HostVulnerabilitySummary>
        {
            new()
            {
                HostName = "PC1",
                Ip = "10.0.0.1",
                TotalVulnCount = 400,
                WindowsVulnCount = 350,
                ProductCount = 4,
                Critical = 1,
                High = 2,
                Medium = 3,
                Low = 4,
                TopProducts = "Windows 11 (300), Chrome (50)"
            }
        };

        var request = new HostVulnerabilityReportRequest("Acme", hosts, "2026-05-23");
        var result = new HostVulnerabilityReportExporter().Export(
            request, dir, "Acme", "2026-05-23_120000", includePdf: true, includeXlsx: true);

        try
        {
            Assert.NotNull(result.PdfPath);
            Assert.NotNull(result.XlsxPath);
            Assert.True(File.Exists(result.PdfPath));
            Assert.True(File.Exists(result.XlsxPath));
            Assert.True(new FileInfo(result.PdfPath!).Length > 500);
            Assert.True(new FileInfo(result.XlsxPath!).Length > 1000);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, true);
        }
    }
}

public class CombinedReportHtmlExporterTests
{
    [Fact]
    public void BuildHtml_IncludesTabsAndCopyHelpers()
    {
        var session = new ReviewSession
        {
            ClientName = "Acme",
            ScanDate = "2026-05-23",
            Presenter = "Tech",
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    Rank = 1,
                    OriginalRank = 1,
                    Product = "Microsoft Edge",
                    RiskScore = 4.5,
                    Epss = 0.3,
                    AvgCvss = 6.1,
                    VulnCount = 12,
                    CveIds = "CVE-2021-33558",
                    RevisedRemediation = "Update Microsoft Edge to the latest version.",
                    IncludeInExport = true,
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC1", Ip = "10.0.0.1", VulnCount = 12 }
                    ]
                }
            ]
        };

        var html = new CombinedReportHtmlExporter(
            new TemplatesService(),
            new RemediationRuleService(),
            new SettingsService()).BuildHtml(session);

        Assert.Contains("Ticket Instructions", html);
        Assert.Contains("Email Template", html);
        Assert.Contains("Ticket Notes (Manage PSA)", html);
        Assert.Contains("copySection('vuln-1')", html);
        Assert.Contains("copyEmailBody()", html);
        Assert.Contains("copyTicketNotes()", html);
        Assert.Contains("PC1", html);
        Assert.DoesNotContain("nvd.nist.gov", html);
    }
}

public class TicketInstructionBuilderTests
{
    private static RemediationRuleService CreateRules() => new();

    [Fact]
    public void BuildSubject_AppliesProductSuffixForRmitPlus()
    {
        var finding = new ReviewFinding { Product = "Microsoft Edge" };
        var subject = TicketInstructionBuilder.BuildSubject(finding, isRmitPlus: true);
        Assert.Contains("Microsoft Edge", subject);
        Assert.Contains("This software updates automatically", subject);
    }

    [Fact]
    public void BuildBodyText_OmitsExcludedHosts()
    {
        var finding = new ReviewFinding
        {
            Product = "Chrome",
            RevisedRemediation = "Update Chrome.",
            AffectedSystems =
            [
                new ReviewAffectedSystem { HostName = "PC1", Ip = "10.0.0.1", VulnCount = 5 },
                new ReviewAffectedSystem { HostName = "PC2", Ip = "10.0.0.2", VulnCount = 3, ExcludedFromExport = true }
            ]
        };

        var body = TicketInstructionBuilder.BuildBodyText(finding, CreateRules());

        Assert.Contains("PC1", body);
        Assert.DoesNotContain("PC2", body);
        Assert.Contains("Affected Systems Count:", body);
        Assert.Contains("1", body);
    }
    [Fact]
    public void BuildSubject_PrependsAfterHoursForRmitPlus()
    {
        var finding = new ReviewFinding
        {
            Product = "Adobe Reader",
            AfterHours = true,
            ThirdParty = true
        };

        var subject = TicketInstructionBuilder.BuildSubject(finding, isRmitPlus: true);

        Assert.StartsWith("After Hours - ", subject);
        Assert.Contains("Adobe Reader", subject);
    }

    [Fact]
    public void BuildSubject_IncludesThirdPartyWhenNotAutoTicket()
    {
        var finding = new ReviewFinding
        {
            Product = "Adobe Reader",
            ThirdParty = true
        };

        var subject = TicketInstructionBuilder.BuildSubject(finding, isRmitPlus: true);

        Assert.Contains("3rd party application", subject);
    }

    [Fact]
    public void BuildBodyText_UsesTicketFormatWhenRemediationUnedited()
    {
        var rules = CreateRules();
        var wordGuidance = rules.GetGuidance("Contoso Legacy Widget", forWord: true);
        var finding = new ReviewFinding
        {
            Product = "Contoso Legacy Widget",
            OriginalRemediation = wordGuidance,
            RevisedRemediation = wordGuidance
        };

        var body = TicketInstructionBuilder.BuildBodyText(finding, rules);

        Assert.Contains("Remediation Instructions:", body);
        Assert.Contains("- Determine device/software identity", body);
        Assert.DoesNotContain("If the client has RMM or scripting available, deploy updates using patch management", body);
    }

    [Fact]
    public void BuildBodyText_IncludesConnectSecureSolutionWhenDistinct()
    {
        var rules = CreateRules();
        var guidance = rules.GetGuidance("Windows 10", forWord: true);
        var finding = new ReviewFinding
        {
            Product = "Windows 10",
            OriginalRemediation = guidance,
            RevisedRemediation = guidance,
            OriginalFix = "Apply Windows Update KB5034123"
        };

        var body = TicketInstructionBuilder.BuildBodyText(finding, rules);

        Assert.Contains("ConnectSecure Solution:", body);
        Assert.Contains("KB5034123", body);
    }
}

public class FindingRemediationExportTests
{
    private static RemediationRuleService CreateRules() => new();

    [Fact]
    public void GetConnectSecureSolution_ReturnsNullWhenFixIsRemediation()
    {
        var fix = "Apply Windows Update KB5034123";
        var finding = new ReviewFinding
        {
            Product = "Windows 10",
            OriginalFix = fix,
            OriginalRemediation = fix,
            RevisedRemediation = fix
        };

        Assert.Null(FindingRemediationExport.GetConnectSecureSolution(finding));
    }

    [Fact]
    public void GetConnectSecureSolution_ReturnsFixWhenRuleGuidanceDiffers()
    {
        var rules = CreateRules();
        var guidance = rules.GetGuidance("Windows 10", forWord: true);
        var finding = new ReviewFinding
        {
            Product = "Windows 10",
            OriginalFix = "Apply Windows Update KB5034123",
            OriginalRemediation = guidance,
            RevisedRemediation = guidance
        };

        Assert.Equal("Apply Windows Update KB5034123", FindingRemediationExport.GetConnectSecureSolution(finding));
    }

    [Fact]
    public void GetTimeEstimateRemediationText_UsesEditedTextWhenChanged()
    {
        var rules = CreateRules();
        var guidance = rules.GetGuidance("Microsoft Edge", forWord: true);
        var finding = new ReviewFinding
        {
            Product = "Microsoft Edge",
            OriginalRemediation = guidance,
            RevisedRemediation = "Client agreed to defer until next quarter."
        };

        var text = FindingRemediationExport.GetTimeEstimateRemediationText(finding, rules);

        Assert.Contains("defer until next quarter", text);
        Assert.DoesNotContain("- Determine device/software identity", text);
    }

    [Fact]
    public void IsRemediationEdited_TrueWhenRevisedDiffers()
    {
        var finding = new ReviewFinding
        {
            OriginalRemediation = "Rule text",
            RevisedRemediation = "Edited in meeting"
        };

        Assert.True(FindingRemediationExport.IsRemediationEdited(finding));
    }
}

public class TimeEstimateModifierHelperTests
{
    [Fact]
    public void IsTicketGenerated_AutoWhenThirdPartyAndAfterHours()
    {
        Assert.True(TimeEstimateModifierHelper.IsTicketGenerated(afterHours: true, ticketGenerated: false, thirdParty: true));
    }

    [Theory]
    [InlineData(false, false, true, "3rd party application")]
    [InlineData(true, false, true, "After-hours ticket generated for 3rd party")]
    public void GetModifierText_CoversKeyCombinations(bool afterHours, bool ticketGenerated, bool thirdParty, string expectedFragment)
    {
        var text = TimeEstimateModifierHelper.GetModifierText(afterHours, ticketGenerated, thirdParty);
        Assert.Contains(expectedFragment, text, StringComparison.OrdinalIgnoreCase);
    }
}

public class FirstPartyVendorHelperTests
{
    [Theory]
    [InlineData("Microsoft Edge", true)]
    [InlineData("HP LaserJet Pro", true)]
    [InlineData("Adobe Reader", false)]
    public void IsFirstPartyVendor_MatchesExpectedProducts(string product, bool expected)
    {
        Assert.Equal(expected, FirstPartyVendorHelper.IsFirstPartyVendor(product));
    }

    [Fact]
    public void IsThirdPartyByDefault_FalseForMicrosoftOnRmitPlus()
    {
        Assert.False(FirstPartyVendorHelper.IsThirdPartyByDefault("Microsoft Edge", isRmitPlus: true));
        Assert.True(FirstPartyVendorHelper.IsThirdPartyByDefault("Adobe Reader", isRmitPlus: true));
    }
}

public class TimeEstimateBuilderTests
{
    [Fact]
    public void Build_IncludesSummaryTotalsForRmitPlus()
    {
        var session = new ReviewSession
        {
            ClientName = "Acme",
            IsRmitPlus = true,
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    OriginalRank = 1,
                    Product = "Adobe Reader",
                    ThirdParty = true,
                    TimeEstimateHours = 2m,
                    IncludeInExport = true,
                    RevisedRemediation = "Update Adobe Reader."
                }
            ]
        };

        var text = TimeEstimateBuilder.Build(session, new RemediationRuleService());

        Assert.Contains("Requires Approval", text);
        Assert.Contains("Total Requiring Approval: 2 hours", text);
        Assert.Contains("Grand Total: 2 hours", text);
    }

    [Fact]
    public void Build_UsesTicketRemediationRules_NotConcatenatedFixes()
    {
        var session = new ReviewSession
        {
            IsRmitPlus = true,
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    OriginalRank = 1,
                    Product = "Microsoft .NET (all versions)",
                    IncludeInExport = true,
                    RevisedRemediation = string.Join("; ", Enumerable.Repeat(
                        "Product is not supported. Reached end of life on 12 Nov 2024", 20))
                }
            ]
        };

        var text = TimeEstimateBuilder.Build(session, new RemediationRuleService());

        Assert.Contains("Two incompatible .NET families", text);
        Assert.DoesNotContain("Reached end of life on 12 Nov 2024", text);
    }

    [Fact]
    public void Build_IncludesUsernamesWithAffectedHostnames()
    {
        var session = new ReviewSession
        {
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    OriginalRank = 1,
                    Product = "Microsoft Teams",
                    IncludeInExport = true,
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "AMI-W11-7", Ip = "10.0.0.7", Username = "jsmith" }
                    ]
                }
            ]
        };

        var text = TimeEstimateBuilder.Build(session, new RemediationRuleService());

        Assert.Contains("Affected Hostnames: AMI-W11-7 (jsmith) - 10.0.0.7", text);
    }
}

public class TicketNotesBuilderTimeEstimateTests
{
    [Fact]
    public void Build_InsertsTicketCreatedLinesWhenGenerated()
    {
        var session = new ReviewSession
        {
            IsRmitPlus = true,
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    OriginalRank = 1,
                    Product = "Adobe Reader",
                    ThirdParty = true,
                    AfterHours = true,
                    IncludeInExport = true
                }
            ]
        };

        var template = new TicketNotesTemplateSettings
        {
            StepsBeforeTickets = "Before",
            StepsAfterTickets = "After",
            ResolvedQuestion = "Q?",
            ResolvedAnswer = "A.",
            NextStepsQuestion = "Next?",
            NextStepsText = "Done."
        };

        var notes = TicketNotesBuilder.Build(session, template);

        Assert.Contains(
            "Ticket created: After Hours - Vulnerability Remediation - Adobe Reader - Update Required",
            notes);
    }

    [Fact]
    public void Build_UsesManageTicketNumberWhenPresent()
    {
        var session = new ReviewSession
        {
            IsRmitPlus = true,
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    OriginalRank = 1,
                    Product = "Adobe Reader",
                    TicketGenerated = true,
                    IncludeInExport = true,
                    ManageTicketNumber = "12345",
                    ManageTicketStatus = "New"
                }
            ]
        };

        var notes = TicketNotesBuilder.Build(session, new TicketNotesTemplateSettings());
        Assert.Contains("- Ticket #12345 (New):", notes);
    }
}

public class CveReferenceHelperTopNTests
{
    [Fact]
    public void IsCveOnlyProduct_DetectsBareCveIds()
    {
        Assert.True(CveReferenceHelper.IsCveOnlyProduct("CVE-2021-33558"));
        Assert.False(CveReferenceHelper.IsCveOnlyProduct("Windows 11"));
        Assert.False(CveReferenceHelper.IsCveOnlyProduct("Google Chrome CVE-2021-33558"));
    }
}

public class CveOnlyFindingDisplayTests
{
    [Fact]
    public void ParseNvdEnrichment_SplitsUrlsAndDescription()
    {
        var parsed = CveOnlyFindingDisplay.ParseNvdEnrichment(
            "https://example.com/advisory | https://vendor.example/patch | Buffer overflow in legacy service.");

        Assert.Equal(2, parsed.ReferenceUrls.Count);
        Assert.Contains("https://example.com/advisory", parsed.ReferenceUrls);
        Assert.Equal("Buffer overflow in legacy service.", parsed.Description);
    }

    [Fact]
    public void GetListSubtitle_UsesNvdDescriptionWhenPresent()
    {
        var finding = new ReviewFinding
        {
            Product = "CVE-2021-33558",
            NvdEnrichment = "https://nvd.nist.gov/x | Remote code execution in example component."
        };

        var subtitle = CveOnlyFindingDisplay.GetListSubtitle(finding);
        Assert.Contains("Remote code execution", subtitle);
    }

    [Fact]
    public void GetListSubtitle_PromptsWhenNvdMissing()
    {
        var finding = new ReviewFinding { Product = "CVE-2021-33558" };
        Assert.Contains("select to load", CveOnlyFindingDisplay.GetListSubtitle(finding), StringComparison.OrdinalIgnoreCase);
    }
}

public class ReviewSessionRankerCveTests
{
    [Fact]
    public void EnsureLegacyFields_KeepsCveFindingsAndNormalizesCveIds()
    {
        var session = new ReviewSession
        {
            ExportTopN = 10,
            Findings =
            [
                new ReviewFinding { OriginalRank = 1, Rank = 1, Product = "Windows 11", IncludeInExport = true, CveIds = "CVE-2024-1" },
                new ReviewFinding { OriginalRank = 2, Rank = 2, Product = "CVE-2021-33558", IncludeInExport = true, CveIds = "" }
            ]
        };

        ReviewSessionRanker.EnsureLegacyFields(session);

        Assert.Equal(2, session.Findings.Count);
        Assert.Equal("CVE-2024-1", session.Findings[0].CveIds);
        Assert.Equal("CVE-2021-33558", session.Findings[1].CveIds);
    }
}

public class OsHostVulnThresholdHelperTests
{
    [Theory]
    [InlineData("Windows 11", OsHostThresholdCategory.Windows11)]
    [InlineData("Windows 10", OsHostThresholdCategory.WindowsOther)]
    [InlineData("Windows Server 2022", OsHostThresholdCategory.WindowsOther)]
    [InlineData("Ubuntu 22.04", OsHostThresholdCategory.OtherOs)]
    [InlineData("Google Chrome", OsHostThresholdCategory.Application)]
    public void ClassifyProduct_DetectsOsCategories(string product, OsHostThresholdCategory expected)
    {
        Assert.Equal(expected, OsHostVulnThresholdHelper.ClassifyProduct(product));
    }

    [Fact]
    public void ShouldIncludeHost_ExcludesBelowWindows11Threshold()
    {
        var settings = new UserSettings { HostVulnThresholdWindows11 = 350 };
        Assert.False(OsHostVulnThresholdHelper.ShouldIncludeHost("Windows 11", "Application", 100, settings));
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Windows 11", "Application", 350, settings));
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Google Chrome", "Application", 1, settings));
    }

    [Fact]
    public void ShouldIncludeHost_IgnoresRegistryAndNetworkFindings()
    {
        var settings = new UserSettings { HostVulnThresholdWindows11 = 350 };
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Windows 11", "Registry", 50, settings));
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Windows 11", "Network", 50, settings));
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("SMB Signing Not Required", "Network", 1, settings));
    }

    [Fact]
    public void ShouldIncludeHost_IgnoresApplicationProductsEvenWhenNamedLikeOs()
    {
        var settings = new UserSettings { HostVulnThresholdWindowsOther = 100 };
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Microsoft Edge", "Application", 1, settings));
        Assert.True(OsHostVulnThresholdHelper.ShouldIncludeHost("Microsoft .NET Runtime", "Application", 1, settings));
    }

    [Fact]
    public void IsOperatingSystemFinding_RequiresApplicationScanSource()
    {
        Assert.True(OsHostVulnThresholdHelper.IsOperatingSystemFinding("Windows 11", "Application"));
        Assert.False(OsHostVulnThresholdHelper.IsOperatingSystemFinding("Windows 11", "Registry"));
        Assert.False(OsHostVulnThresholdHelper.IsOperatingSystemFinding("Google Chrome", "Application"));
    }

    [Fact]
    public void HostOsThresholdApplier_ExcludesLowCountHostsForOsFindings()
    {
        var session = new ReviewSession
        {
            Findings =
            [
                new ReviewFinding
                {
                    Product = "Windows 11",
                    Source = "Application",
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC1", VulnCount = 400 },
                        new ReviewAffectedSystem { HostName = "PC2", VulnCount = 50 }
                    ]
                },
                new ReviewFinding
                {
                    Product = "Windows 11",
                    Source = "Registry",
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC3", VulnCount = 50 }
                    ]
                },
                new ReviewFinding
                {
                    Product = "Google Chrome",
                    Source = "Application",
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC4", VulnCount = 1 }
                    ]
                }
            ]
        };

        HostOsThresholdApplier.Apply(session, new UserSettings { HostVulnThresholdWindows11 = 350 });

        Assert.False(session.Findings[0].AffectedSystems[0].ExcludedFromExport);
        Assert.True(session.Findings[0].AffectedSystems[1].ExcludedFromExport);
        Assert.False(session.Findings[1].AffectedSystems[0].ExcludedFromExport);
        Assert.False(session.Findings[2].AffectedSystems[0].ExcludedFromExport);
    }
}

public class ProductTypeSuffixHelperTests
{
    [Theory]
    [InlineData("Windows 11", false, "Updates Required")]
    [InlineData("Microsoft Edge", true, "This software updates automatically")]
    [InlineData("Microsoft Edge", false, "This software updates automatically")]
    [InlineData("Google Chrome", false, "This software updates automatically")]
    public void GetSuffix_MatchesPowerShellRules(string product, bool isRmitPlus, string expectedFragment)
    {
        var suffix = ProductTypeSuffixHelper.GetSuffix(product, isRmitPlus);
        Assert.Contains(expectedFragment, suffix);
    }
}

public class EmailTemplateBuilderTests
{
    [Fact]
    public void Build_UsesApprovalNoteForRmitClients()
    {
        var session = new ReviewSession { ClientName = "Acme", Presenter = "Tech", ExportTopN = 10, IsRmitPlus = false };
        var template = new EmailTemplateSettings
        {
            SubjectFormat = "{Year} Q{Quarter} Follow Up",
            Body = "Hello {Greeting}\n\n{NoteText}\n\n{PreparedBy}"
        };

        var email = EmailTemplateBuilder.Build(session, template);
        Assert.Contains("No remediation will begin without your approval", email);
        Assert.DoesNotContain("RMIT+ agreement", email);
    }

    [Fact]
    public void Build_UsesRmitPlusNoteWhenEnabled()
    {
        var session = new ReviewSession { ClientName = "Acme", Presenter = "Tech", ExportTopN = 10, IsRmitPlus = true };
        var template = new EmailTemplateSettings
        {
            SubjectFormat = "{Year} Q{Quarter} Follow Up",
            Body = "Hello {Greeting}\n\n{NoteText}\n\n{PreparedBy}"
        };

        var email = EmailTemplateBuilder.Build(session, template);
        Assert.Contains("RMIT+ agreement", email);
    }

    [Fact]
    public void Build_SelectsRmitPlusTemplateWhenSessionIsRmitPlus()
    {
        var session = new ReviewSession { ClientName = "Acme", Presenter = "Tech", ExportTopN = 10, IsRmitPlus = true };
        var templates = new VScanMagicTemplates
        {
            EmailTemplate = new EmailTemplateSettings { Body = "RMIT body: {NoteText}" },
            EmailTemplateRmitPlus = new EmailTemplateSettings { Body = "RMIT+ body: {NoteText}" }
        };

        var email = EmailTemplateBuilder.Build(session, templates);
        Assert.Contains("RMIT+ body:", email);
        Assert.Contains("RMIT+ agreement", email);
    }

    [Fact]
    public void Build_SelectsRmitTemplateWhenSessionIsNotRmitPlus()
    {
        var session = new ReviewSession { ClientName = "Acme", Presenter = "Tech", ExportTopN = 10, IsRmitPlus = false };
        var templates = new VScanMagicTemplates
        {
            EmailTemplate = new EmailTemplateSettings { Body = "RMIT body: {NoteText}" },
            EmailTemplateRmitPlus = new EmailTemplateSettings { Body = "RMIT+ body: {NoteText}" }
        };

        var email = EmailTemplateBuilder.Build(session, templates);
        Assert.Contains("RMIT body:", email);
        Assert.Contains("No remediation will begin without your approval", email);
    }

    [Fact]
    public void Build_SubstitutesDeliverableLinks()
    {
        var session = new ReviewSession
        {
            ClientName = "Acme",
            Presenter = "Tech",
            ExportTopN = 10,
            TopNReportUrl = "https://sharepoint/topn",
            ReportsFolderUrl = "https://sharepoint/folder",
            SchedulingLinkUrl = "https://timezest/meet"
        };
        var template = new EmailTemplateSettings
        {
            Body = "Top: {TopNReportLink}\nFolder: {ReportsFolderLink}\nSchedule: {SchedulingLink}"
        };
        var links = new DeliverableLinks
        {
            TopNReportUrl = session.TopNReportUrl,
            ReportsFolderUrl = session.ReportsFolderUrl,
            SchedulingLinkUrl = session.SchedulingLinkUrl
        };

        var email = EmailTemplateBuilder.Build(session, template, isRmitPlus: false, links);

        Assert.Contains("https://sharepoint/topn", email);
        Assert.Contains("https://sharepoint/folder", email);
        Assert.Contains("https://timezest/meet", email);
    }

    [Fact]
    public void BuildHtmlBody_EmbedsAnchorsForDeliverableLinks()
    {
        var session = new ReviewSession
        {
            ClientName = "Acme",
            Presenter = "Tech",
            ExportTopN = 10,
            Findings = Enumerable.Range(1, 10)
                .Select(i => new ReviewFinding
                {
                    Source = "Application",
                    Product = $"App {i}",
                    IncludeInExport = true,
                    OriginalRank = i,
                    Rank = i
                })
                .ToList()
        };
        var links = new DeliverableLinks
        {
            TopNReportUrl = "https://sharepoint/topn",
            ReportsFolderUrl = "https://sharepoint/folder",
            SchedulingLinkUrl = "https://timezest/meet"
        };
        var body = """
Recommended remediation priorities (Top Ten):
https://sharepoint/topn
Complete report package:
https://sharepoint/folder
Schedule time with me
https://timezest/meet
""";

        var (plain, html) = EmailTemplateBuilder.PrepareDeliverableCopy(body, links, "Top Ten");

        Assert.Contains("Open Top Ten report", plain);
        Assert.Contains("Open complete report package", plain);
        Assert.Contains("Schedule Time With Me", plain);
        Assert.DoesNotContain("https://sharepoint/topn", plain);
        Assert.Contains("<a href=\"https://sharepoint/topn\">Open Top Ten report</a>", html);
        Assert.Contains("<a href=\"https://sharepoint/folder\">Open complete report package</a>", html);
        Assert.Contains("<a href=\"https://timezest/meet\">Schedule Time With Me</a>", html);
        Assert.Contains("<p style=\"margin:0 0 12px 0;\"><a href=\"https://sharepoint/topn\">Open Top Ten report</a></p>", html);
        Assert.Contains("<p style=\"margin:0 0 12px 0;\">Complete report package:</p>", html);
    }

    [Fact]
    public void NormalizeDeliverableBodySpacing_MatchesExpectedLayout()
    {
        var body = """
Good morning,

Your quarterly vulnerability scan report has been completed and is available in your client folder.

Recommended remediation priorities (Top Ten):
https://sharepoint/topn
Complete report package:
https://sharepoint/folder
The folder contains the following reports:
• Pending Remediation EPSS Score Report – classifies vulnerabilities.
• All Vulnerabilities Report – a comprehensive list.
Not all vulnerabilities may be feasible to remediate depending on business or technical constraints.

Schedule time with me
https://timezest/meet
Note: No remediation will begin without your approval.
""";

        var normalized = EmailTemplateBuilder.NormalizeDeliverableBodySpacing(body);

        Assert.Contains("https://sharepoint/topn\n\nComplete report package:", normalized);
        Assert.Contains("https://sharepoint/folder\n\nThe folder contains", normalized);
        Assert.Contains("• Pending Remediation EPSS Score Report – classifies vulnerabilities.\n• All Vulnerabilities", normalized);
        Assert.DoesNotContain("• Pending Remediation EPSS Score Report – classifies vulnerabilities.\n\n• All Vulnerabilities", normalized);
        Assert.Contains("https://timezest/meet\n\nNote:", normalized);
        Assert.DoesNotContain("Schedule time with me\nhttps://timezest/meet", normalized);
    }

    [Fact]
    public void NormalizeDeliverableBodySpacing_CollapsesBlanksBetweenLaterBullets()
    {
        var body = """
The folder contains the following reports:
• Pending Remediation EPSS Score Report – classifies vulnerabilities.

• All Vulnerabilities Report – a comprehensive list.

• Executive Summary Report – a high-level overview.

• External Scan – detected vulnerabilities.
""";

        var normalized = EmailTemplateBuilder.NormalizeDeliverableBodySpacing(body);

        Assert.Contains("• Pending Remediation EPSS Score Report – classifies vulnerabilities.\n• All Vulnerabilities Report", normalized);
        Assert.DoesNotContain("• All Vulnerabilities Report – a comprehensive list.\n\n• Executive Summary Report", normalized);
        Assert.Contains("• Executive Summary Report – a high-level overview.\n• External Scan", normalized);
    }

    [Fact]
    public void BuildHtmlBody_DoesNotCorruptAnchorsWhenFolderUrlIsPrefixOfTopNUrl()
    {
        var links = new DeliverableLinks
        {
            TopNReportUrl = "https://sharepoint/sites/client/Reports/Top13.docx",
            ReportsFolderUrl = "https://sharepoint/sites/client/Reports",
            SchedulingLinkUrl = "https://timezest/meet"
        };
        var body = """
Recommended remediation priorities (Top 13):
https://sharepoint/sites/client/Reports/Top13.docx
Complete report package:
https://sharepoint/sites/client/Reports
Schedule time with me
https://timezest/meet
Note: No remediation will begin without your approval.
""";

        var html = EmailTemplateBuilder.BuildHtmlBody(body, links, "Top 13");

        Assert.Contains("<a href=\"https://sharepoint/sites/client/Reports/Top13.docx\">Open Top 13 report</a>", html);
        Assert.Contains("<a href=\"https://sharepoint/sites/client/Reports\">Open complete report package</a>", html);
        Assert.Contains("<a href=\"https://timezest/meet\">Schedule Time With Me</a>", html);
        Assert.DoesNotContain("Schedule time with me\">Open Top 13 report", html);
        Assert.Contains("Open Top 13 report</a></p>", html);
        Assert.Contains("&nbsp;</p>", html);
    }

    [Fact]
    public void PrepareDeliverableCopy_UsesSchedulingLabelWhenTopNAndSchedulingUrlsMatch()
    {
        var sharedUrl = "https://sharepoint/sites/client/Reports/Top13.docx";
        var links = new DeliverableLinks
        {
            TopNReportUrl = sharedUrl,
            ReportsFolderUrl = "https://sharepoint/sites/client/Reports",
            SchedulingLinkUrl = sharedUrl
        };
        var body = """
Recommended remediation priorities (Top 13):
https://sharepoint/sites/client/Reports/Top13.docx
Complete report package:
https://sharepoint/sites/client/Reports
Not all vulnerabilities may be feasible to remediate depending on business or technical constraints.
https://sharepoint/sites/client/Reports/Top13.docx
Note: No remediation will begin without your approval.
""";

        var (plain, html) = EmailTemplateBuilder.PrepareDeliverableCopy(body, links, "Top 13");

        Assert.Contains("Open Top 13 report", plain);
        Assert.Contains("Schedule Time With Me", plain);
        Assert.Contains("<a href=\"https://sharepoint/sites/client/Reports/Top13.docx\">Open Top 13 report</a>", html);
        Assert.Contains("<a href=\"https://sharepoint/sites/client/Reports/Top13.docx\">Schedule Time With Me</a>", html);
    }

    [Fact]
    public void NormalizeDeliverableBodySpacing_RemovesDuplicateSchedulingHeader()
    {
        var body = """
Not all vulnerabilities may be feasible to remediate depending on business or technical constraints.

Schedule time with me

https://timezest/meet

Note: No remediation will begin without your approval.
""";

        var normalized = EmailTemplateBuilder.NormalizeDeliverableBodySpacing(body);

        Assert.Contains("constraints.\n\nhttps://timezest/meet\n\nNote:", normalized);
        Assert.DoesNotContain("Schedule time with me\n\nhttps://timezest/meet", normalized);
    }
}

public class ReviewUsernameRefreshServiceTests
{
    [Fact]
    public void CollectHostnamesNeedingUsernames_OnlyExportFindingsWhenRequested()
    {
        var session = new ReviewSession
        {
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    Product = "Chrome",
                    Source = "Application",
                    IncludeInExport = true,
                    OriginalRank = 1,
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC1", Username = "" }
                    ]
                },
                new ReviewFinding
                {
                    Product = "Reserve App",
                    Source = "Application",
                    IncludeInExport = false,
                    OriginalRank = 2,
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC2", Username = "" }
                    ]
                }
            ]
        };

        var exportOnly = ReviewUsernameRefreshService.CollectHostnamesNeedingUsernames(session, exportFindingsOnly: true);
        var all = ReviewUsernameRefreshService.CollectHostnamesNeedingUsernames(session, exportFindingsOnly: false);

        Assert.Equal(["PC1"], exportOnly);
        Assert.Equal(["PC1", "PC2"], all);
    }

    [Fact]
    public void CollectHostnamesNeedingUsernames_SkipsHostsThatAlreadyHaveUsernames()
    {
        var session = new ReviewSession
        {
            Findings =
            [
                new ReviewFinding
                {
                    Product = "Chrome",
                    Source = "Application",
                    IncludeInExport = true,
                    AffectedSystems =
                    [
                        new ReviewAffectedSystem { HostName = "PC1", Username = "jsmith" },
                        new ReviewAffectedSystem { HostName = "PC2", Username = "" }
                    ]
                }
            ]
        };

        var hostnames = ReviewUsernameRefreshService.CollectHostnamesNeedingUsernames(session);

        Assert.Equal(["PC2"], hostnames);
    }
}

public class CveReferenceHelperTests
{
    [Fact]
    public void SplitCveIds_ExtractsUniqueIds()
    {
        var ids = CveReferenceHelper.SplitCveIds("CVE-2021-33558; CVE-2020-15778; cve-2021-33558");
        Assert.Equal(["CVE-2021-33558", "CVE-2020-15778"], ids);
    }

    [Fact]
    public void GetNvdDetailUrl_UsesUppercaseCveId()
    {
        Assert.Equal("https://nvd.nist.gov/vuln/detail/CVE-2021-33558",
            CveReferenceHelper.GetNvdDetailUrl("cve-2021-33558"));
    }

    [Fact]
    public void MergeCveIds_DeduplicatesValues()
    {
        Assert.Equal("CVE-2021-33558; CVE-2020-15778",
            CveReferenceHelper.MergeCveIds("CVE-2021-33558", "CVE-2020-15778; CVE-2021-33558"));
    }
}

public class NvdRemediationFormatterTests
{
    [Fact]
    public void BuildSummary_PrefersAdvisoryLinksAndDescription()
    {
        const string json = """
            {
              "cve": {
                "descriptions": [{ "lang": "en", "value": "Example firmware vulnerability." }],
                "references": [
                  { "url": "https://vendor.example/security/advisory-123", "source": "vendor" },
                  { "url": "https://example.com/other", "source": "misc" }
                ]
              }
            }
            """;

        using var document = System.Text.Json.JsonDocument.Parse(json);
        var summary = NvdRemediationFormatter.BuildSummary(document.RootElement);

        Assert.Contains("vendor.example/security/advisory-123", summary);
        Assert.Contains("Example firmware vulnerability.", summary);
    }
}

public class CveEnrichmentPolicyTests
{
    [Fact]
    public void IngestEnrichmentTargets_IncludeOnlyExportFindings()
    {
        var session = new ReviewSession
        {
            Findings =
            [
                new ReviewFinding { Product = "Chrome", IncludeInExport = true },
                new ReviewFinding { Product = "CVE-1", Source = "Registry", IncludeInExport = false }
            ]
        };

        var targets = session.Findings.Where(f => f.IncludeInExport).ToList();
        Assert.Single(targets);
        Assert.Equal("Chrome", targets[0].Product);
    }

    [Fact]
    public void AppendNvdContext_AddsSummaryWhenFixMissing()
    {
        var finding = new ReviewFinding
        {
            Product = "CVE-2021-33558",
            CveIds = "CVE-2021-33558",
            NvdEnrichment = "https://vendor.example/advisory | Example firmware vulnerability."
        };

        var text = CveEnrichmentPolicy.AppendNvdContext("Review vendor advisory.", finding);

        Assert.Contains("Review vendor advisory.", text);
        Assert.Contains("NVD / advisory context:", text);
        Assert.Contains("vendor.example/advisory", text);
    }

    [Fact]
    public void GetTicketRemediationText_CveOnly_UsesMinimalStepsNotGenericCatchAll()
    {
        var finding = new ReviewFinding
        {
            Product = "CVE-2021-33558",
            CveIds = "CVE-2021-33558",
            OriginalRemediation = "Catch-all from ingest",
            RevisedRemediation = "Catch-all from ingest"
        };

        var text = FindingRemediationExport.GetTicketRemediationText(finding, new RemediationRuleService());

        Assert.Contains("NVD link in CVE references", text);
        Assert.DoesNotContain("Determine device/software identity (Product/OS, affected hosts)", text);
    }

    [Fact]
    public void GetTicketRemediationText_CveOnly_RespectsEditedRemediation()
    {
        var finding = new ReviewFinding
        {
            Product = "CVE-2021-33558",
            CveIds = "CVE-2021-33558",
            OriginalRemediation = "Original",
            RevisedRemediation = "Technician agreed: patch firmware on PREVIANT-S6D16."
        };

        var text = FindingRemediationExport.GetTicketRemediationText(finding, new RemediationRuleService());

        Assert.Contains("Technician agreed", text);
        Assert.DoesNotContain("Determine device/software identity", text);
    }
}

public class CveExportFormatterTests
{
    [Fact]
    public void UsesCveExportTreatment_OnlyForCveOnlyProductWithIds()
    {
        Assert.True(CveExportFormatter.UsesCveExportTreatment(new ReviewFinding
        {
            Product = "CVE-2015-0240",
            CveIds = "CVE-2015-0240"
        }));
        Assert.False(CveExportFormatter.UsesCveExportTreatment(new ReviewFinding
        {
            Product = "Google Chrome",
            CveIds = "CVE-2024-1234"
        }));
    }

    [Fact]
    public void FormatReferencesSection_IncludesNvdUrls()
    {
        var section = CveExportFormatter.FormatReferencesSection(new ReviewFinding
        {
            Product = "CVE-2015-0240",
            CveIds = "CVE-2015-0240"
        });

        Assert.Contains("CVE-2015-0240", section);
        Assert.Contains("https://nvd.nist.gov/vuln/detail/CVE-2015-0240", section);
    }

    [Fact]
    public void TicketBody_CveOnly_IncludesReferencesAndOmitsUpdateRequiredSuffix()
    {
        var finding = new ReviewFinding
        {
            Product = "CVE-2015-0240",
            CveIds = "CVE-2015-0240",
            RiskScore = 20.06,
            Epss = 0.91,
            AvgCvss = 7,
            VulnCount = 1,
            OriginalRemediation = "x",
            RevisedRemediation = "x",
            AffectedSystems =
            [
                new ReviewAffectedSystem { HostName = "PREVIANT-S6D16", Ip = "192.168.0.52", VulnCount = 1 }
            ]
        };

        var subject = FindingTitleFormatter.FormatTicketSubject(finding, isRmitPlus: false);
        var body = TicketInstructionBuilder.BuildBodyText(finding, new RemediationRuleService());

        Assert.Equal("Vulnerability Remediation - CVE-2015-0240 - Investigate and Resolve", subject);
        Assert.Contains("CVE references:", body);
        Assert.Contains("https://nvd.nist.gov/vuln/detail/CVE-2015-0240", body);
        Assert.DoesNotContain(" - Update Required", subject);
        Assert.Contains("Investigate and Resolve", subject);
        Assert.DoesNotContain("Determine device/software identity (Product/OS, affected hosts)", body);
        Assert.Contains("NVD link in CVE references", body);
    }

    [Fact]
    public void TicketNotes_RmitPlus_ListsFullSubjectWithCveSuffix()
    {
        var session = new ReviewSession
        {
            IsRmitPlus = true,
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    Product = "CVE-2015-0240",
                    CveIds = "CVE-2015-0240",
                    IncludeInExport = true,
                    TicketGenerated = true,
                    OriginalRank = 1,
                    Rank = 1
                }
            ]
        };

        var notes = TicketNotesBuilder.Build(session, TicketNotesTemplateSettings.CreateDefault(), isRmitPlus: true);

        Assert.Contains("Ticket created: Vulnerability Remediation - CVE-2015-0240 - Investigate and Resolve", notes);
        Assert.DoesNotContain(" - Update Required", notes);
    }

    [Fact]
    public void TicketBody_NamedProductWithCve_DoesNotUseCveReferencesSection()
    {
        var finding = new ReviewFinding
        {
            Product = "Some Third Party App",
            CveIds = "CVE-2024-1234",
            RiskScore = 10,
            Epss = 0.5,
            VulnCount = 3,
            OriginalRemediation = "Update",
            RevisedRemediation = "Update"
        };

        var body = TicketInstructionBuilder.BuildBodyText(finding, new RemediationRuleService());

        Assert.DoesNotContain("CVE references:", body);
        Assert.Contains(" - Update Required", FindingTitleFormatter.FormatTicketSubject(finding, false));
    }
}

public class FindingExportDetailsTests
{
    [Fact]
    public void IncludedSystems_OmitsExcludedHosts()
    {
        var finding = new ReviewFinding
        {
            AffectedSystems =
            [
                new ReviewAffectedSystem { HostName = "PC1", Ip = "10.0.0.1", VulnCount = 5 },
                new ReviewAffectedSystem { HostName = "PC2", Ip = "10.0.0.2", VulnCount = 3, ExcludedFromExport = true }
            ]
        };

        var included = FindingExportDetails.IncludedSystems(finding);
        Assert.Single(included);
        Assert.Equal("PC1", included[0].HostName);
    }

    [Fact]
    public void GetIncludedVulnCount_SumsIncludedHostsOnly()
    {
        var finding = new ReviewFinding
        {
            VulnCount = 100,
            AffectedSystems =
            [
                new ReviewAffectedSystem { HostName = "PC1", Ip = "10.0.0.1", VulnCount = 5 },
                new ReviewAffectedSystem { HostName = "PC2", Ip = "10.0.0.2", VulnCount = 3, ExcludedFromExport = true },
                new ReviewAffectedSystem { HostName = "PC3", Ip = "10.0.0.3", VulnCount = 7 }
            ]
        };

        Assert.Equal(12, FindingExportDetails.GetIncludedVulnCount(finding));
    }

    [Fact]
    public void FormatAffectedSystemCompact_IncludesUsernameAndIpLikePowerShell()
    {
        var text = FindingExportDetails.FormatAffectedSystemCompact(new ReviewAffectedSystem
        {
            HostName = "AMI-W11-7",
            Ip = "10.0.0.7",
            Username = "jsmith"
        });

        Assert.Equal("AMI-W11-7 (jsmith) - 10.0.0.7", text);
    }

    [Fact]
    public void FormatAffectedSystemsCompactInline_JoinsHostsWithUsernames()
    {
        var finding = new ReviewFinding
        {
            AffectedSystems =
            [
                new ReviewAffectedSystem { HostName = "AMI-W11-7", Ip = "10.0.0.7", Username = "jsmith" },
                new ReviewAffectedSystem { HostName = "AMI-W11-6", Ip = "10.0.0.6", Username = "alee" }
            ]
        };

        var text = FindingExportDetails.FormatAffectedSystemsCompactInline(finding);

        Assert.Equal("AMI-W11-7 (jsmith) - 10.0.0.7, AMI-W11-6 (alee) - 10.0.0.6", text);
    }

    [Fact]
    public void FormatReferenceLinks_JoinsCveIds()
    {
        var text = CveReferenceHelper.FormatReferenceLinks("CVE-2021-33558; CVE-2020-15778", "; ");
        Assert.Equal("CVE-2021-33558; CVE-2020-15778", text);
    }
}

public class ConnectSecureFixFormatterTests
{
    [Theory]
    [InlineData("['5077181']", "Apply Windows Update KB5077181")]
    [InlineData("[\"148.0.0\"]", "Update to version 148.0.0 or later")]
    [InlineData("['5077181']; ['148.0.0']", "Apply Windows Update KB5077181. Update to version 148.0.0 or later")]
    [InlineData("[\"Microsoft Edge 131.0.2903.99\"]", "Microsoft Edge 131.0")]
    [InlineData("None", "")]
    [InlineData("[\"None\"]", "")]
    public void ToReadableText_FormatsPatchAndVersionValues(string input, string expected)
    {
        Assert.Equal(expected, ConnectSecureFixFormatter.ToReadableText(input));
    }

    [Fact]
    public void EnsureLegacyFields_CleansBracketedFixText()
    {
        var session = new ReviewSession { ExportTopN = 1 };
        session.Findings.Add(new ReviewFinding
        {
            OriginalRank = 1,
            Rank = 1,
            Product = "Microsoft Edge",
            OriginalFix = "['5077181']",
            IncludeInExport = true
        });

        ReviewSessionRanker.EnsureLegacyFields(session);

        Assert.Equal("Apply Windows Update KB5077181", session.Findings[0].OriginalFix);
    }
}

public class ProductDisplayFormatterTests
{
    [Fact]
    public void GetProductMajorVersion_TrimsPatchLevels()
    {
        Assert.Equal("MongoDB 3.4", ProductConsolidator.GetProductMajorVersion("MongoDB 3.4.24"));
        Assert.Equal("Adobe Reader", ProductConsolidator.GetProductMajorVersion("Adobe Reader 11.0"));
    }

    [Theory]
    [InlineData("Microsoft .NET Runtime 8.0.14 (x64)", "Microsoft .NET (all versions)")]
    [InlineData("Microsoft .NET Framework 4.8", "Microsoft .NET (all versions)")]
    [InlineData("Microsoft .NET Host - 8.0.14 (x64)", "Microsoft .NET (all versions)")]
    [InlineData("Microsoft .NET SDK 8.0.100", "Microsoft .NET (all versions)")]
    [InlineData("Microsoft Visual C++ 2015-2022 Redistributable", "Microsoft Visual C++ (all versions)")]
    [InlineData("Google Chrome", "Google Chrome")]
    public void GetTimeEstimateGroupKey_GroupsDotNetAndVisualCppVariants(string input, string expected) =>
        Assert.Equal(expected, ProductConsolidator.GetTimeEstimateGroupKey(input));
}

public class RemediationRuleGuidanceTests
{
    private static RemediationRule GetDefaultRule(string pattern) =>
        RemediationRuleDefaults.GetAll().First(r => r.Pattern == pattern);

    [Fact]
    public void DefaultRules_VisualCpp_ExplainsSideBySideYearLockIn()
    {
        var rule = GetDefaultRule("*Visual C++*");

        Assert.Contains("side-by-side", rule.WordText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("major release lines", rule.WordText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2015–2022", rule.WordText);
        Assert.Contains("not interchangeable", rule.TicketText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DefaultRules_DotNetAllVersions_ExplainsFrameworkVsModernBreak()
    {
        var rule = GetDefaultRule("*Microsoft .NET (all versions)*");

        Assert.Contains(".NET Framework", rule.WordText);
        Assert.Contains("modern .NET", rule.WordText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".NET 5", rule.WordText);
        Assert.Contains("not backward compatible", rule.WordText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DefaultRules_Chrome_UsesAutoUpdateStyleAndVerifyFirstGuidance()
    {
        var rule = GetDefaultRule("*Google Chrome*");
        Assert.Equal(RemediationGuidanceStyle.AutoUpdate, rule.GuidanceStyle);
        Assert.Contains("Google Update", rule.WordText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("verify first", rule.TicketText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DefaultRules_Tls10_IsConfigurationGuidance()
    {
        var rule = GetDefaultRule("*TLSv1.0*");
        Assert.Contains("TLS 1.0", rule.WordText);
        Assert.Contains("disable", rule.WordText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LoadRules_UpgradesStalePersistedDefaultsWhenRevisionIncreases()
    {
        var configDir = Path.Combine(Path.GetTempPath(), "VScanMagicTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(configDir);

        try
        {
            var path = VScanMagicPaths.RemediationRulesFile(configDir);
            var staleRules = RemediationRuleDefaults.GetAll()
                .Select(r => new RemediationRule
                {
                    Pattern = r.Pattern,
                    WordText = r.Pattern == "*Google Chrome*" ? "Old Chrome guidance." : r.WordText,
                    TicketText = r.Pattern == "*Dot Net 6*" ? "- Patch .NET 6 runtime via Windows Update until migrated" : r.TicketText,
                    IsDefault = r.IsDefault,
                    GuidanceStyle = r.GuidanceStyle
                })
                .ToList();

            File.WriteAllText(path, JsonSerializer.Serialize(staleRules));

            var service = new RemediationRuleService(configDir);
            var rules = service.LoadRules();

            var chrome = rules.First(r => r.Pattern == "*Google Chrome*");
            Assert.Contains("Google Update", chrome.WordText, StringComparison.OrdinalIgnoreCase);

            var dotNet6 = rules.First(r => r.Pattern == "*Dot Net 6*");
            Assert.Contains("November 12, 2024", dotNet6.WordText);
            Assert.DoesNotContain("Patch .NET 6 runtime via Windows Update", dotNet6.TicketText);

            Assert.Equal(RemediationRuleDefaults.Revision.ToString(),
                File.ReadAllText(Path.ChangeExtension(path, ".defaults_revision")).Trim());
        }
        finally
        {
            if (Directory.Exists(configDir))
                Directory.Delete(configDir, recursive: true);
        }
    }
}

public class ProductNameNormalizerTests
{
    [Theory]
    [InlineData("[\"Microsoft Edge\"]", "Microsoft Edge")]
    [InlineData("[\"Google Chrome\", \"Microsoft Edge\"]", "Google Chrome, Microsoft Edge")]
    [InlineData("[\"Microsoft 365 Apps for business - en-us\"]", "Microsoft 365 Apps for business - en-us")]
    [InlineData("[\"CVE-2021-33558\"]", "CVE-2021-33558")]
    [InlineData("Windows 11 (all versions)", "Windows 11 (all versions)")]
    public void Normalize_RemovesJsonArrayBrackets(string input, string expected)
    {
        Assert.Equal(expected, ProductNameNormalizer.Normalize(input));
    }

    [Fact]
    public void ParseProductNames_SplitsJsonArrayIntoParts()
    {
        var parts = ProductNameNormalizer.ParseProductNames("[\"Microsoft Edge\", \"Google Chrome\"]");
        Assert.Equal(["Microsoft Edge", "Google Chrome"], parts);
    }

    [Fact]
    public void FormatDisplayName_TrimsPatchVersions()
    {
        Assert.Equal("MongoDB 3.4", ProductNameNormalizer.FormatDisplayName("[\"MongoDB 3.4.24\"]"));
    }

    [Fact]
    public void EnsureLegacyFields_CleansBracketedProductNames()
    {
        var session = new ReviewSession { ExportTopN = 1 };
        session.Findings.Add(new ReviewFinding
        {
            OriginalRank = 1,
            Rank = 1,
            Product = "[\"Microsoft Edge\"]",
            IncludeInExport = true
        });

        ReviewSessionRanker.EnsureLegacyFields(session);

        Assert.Equal("Microsoft Edge", session.Findings[0].Product);
    }
}

public class HostVulnerabilitySummarizerTests
{
    [Fact]
    public void Summarize_GroupsRecordsByHost()
    {
        var records = new List<VulnerabilityRecord>
        {
            new() { HostName = "PC1", Ip = "10.0.0.1", Product = "Windows 11", Critical = 1, VulnerabilityCount = 1 },
            new() { HostName = "PC1", Ip = "10.0.0.1", Product = "Chrome", High = 2, VulnerabilityCount = 2 },
            new() { HostName = "PC2", Ip = "10.0.0.2", Product = "Windows 10", High = 5, VulnerabilityCount = 5 }
        };

        var hosts = new HostVulnerabilitySummarizer().Summarize(records);
        Assert.Equal(2, hosts.Count);
        var pc1 = hosts.Single(h => h.HostName == "PC1");
        Assert.Equal(3, pc1.TotalVulnCount);
        Assert.Equal(1, pc1.WindowsVulnCount);
        Assert.Equal(2, pc1.ProductCount);
    }
}

public class ReportArchiveHelperTests
{
    [Fact]
    public void EnsureExtractedReportFile_UnwrapsNestedXlsxWrapper()
    {
        var wrapper = Path.Combine(
            Path.GetTempPath(),
            "VScanMagic",
            "Downloads",
            "Fabrikam Industries Inc - All Vulnerabilities Report_2026-05-23_15-53-06.xlsx");
        if (!File.Exists(wrapper)) return;

        var extracted = Core.IO.ReportArchiveHelper.EnsureExtractedReportFile(wrapper);
        try
        {
            Assert.NotEqual(wrapper, extracted);
            using var archive = System.IO.Compression.ZipFile.OpenRead(extracted);
            Assert.Contains(archive.Entries, e =>
                e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            if (extracted != wrapper && File.Exists(extracted))
                File.Delete(extracted);
        }
    }

    [Fact]
    public void NormalizeDownloadedReportFile_ReplacesWrapperZipAndRemovesSidecars()
    {
        var dir = Path.Combine(Path.GetTempPath(), "VScanMagicTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var target = Path.Combine(dir, "Client - Executive Summary Report_2026-05-28_12-00-00.docx");
        var innerDocx = CreateMinimalDocx(Path.Combine(dir, "inner.docx"));

        try
        {
            CreateWrapperZip(target, innerDocx, "Executive Summary Report.docx");
            Assert.True(File.Exists(target));

            Core.IO.ReportArchiveHelper.NormalizeDownloadedReportFile(target);

            Assert.True(File.Exists(target));
            Assert.False(File.Exists(target + ".extracting"));
            Assert.False(File.Exists(target + ".extracted"));
            using var archive = System.IO.Compression.ZipFile.OpenRead(target);
            Assert.Contains(archive.Entries, e =>
                e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            try
            {
                if (Directory.Exists(dir))
                    Directory.Delete(dir, recursive: true);
            }
            catch
            {
                // Best-effort cleanup for temp test dirs.
            }
        }
    }

    private static string CreateMinimalDocx(string path)
    {
        using var archive = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create);
        archive.CreateEntry("[Content_Types].xml");
        return path;
    }

    private static void CreateWrapperZip(string wrapperPath, string innerFilePath, string innerEntryName)
    {
        if (File.Exists(wrapperPath))
            File.Delete(wrapperPath);

        using var archive = System.IO.Compression.ZipFile.Open(wrapperPath, System.IO.Compression.ZipArchiveMode.Create);
        archive.CreateEntryFromFile(innerFilePath, innerEntryName, System.IO.Compression.CompressionLevel.Optimal);
    }
}

public class ExcelCellTextTests
{
    [Fact]
    public void Truncate_LimitsToExcelMaximum()
    {
        var text = new string('x', ExcelCellText.MaxLength + 100);
        var truncated = ExcelCellText.Truncate(text);
        Assert.True(truncated.Length <= ExcelCellText.MaxLength);
        Assert.EndsWith("…", truncated);
    }

    [Fact]
    public void FormatAffectedSystemsForExcel_SummarizesLargeHostLists()
    {
        var finding = new ReviewFinding
        {
            Product = "Windows 11",
            AffectedSystems = Enumerable.Range(1, 5000)
                .Select(i => new ReviewAffectedSystem
                {
                    HostName = $"WORKSTATION-{i:D4}",
                    Ip = $"10.0.{i / 256}.{i % 256}",
                    VulnCount = 3
                })
                .ToList()
        };

        var text = ExcelCellText.FormatAffectedSystemsForExcel(finding);
        Assert.True(text.Length <= ExcelCellText.MaxLength);
        Assert.Contains("5000 hosts total", text);
        Assert.Contains("Affected Systems sheet", text);
    }

    [Fact]
    public void FlatXlsxExporter_HandlesLargeAffectedHostLists()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vscanmagic_xlsx_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "export.xlsx");

        var session = new ReviewSession
        {
            ClientName = "Acme",
            ScanDate = "2026-05-23",
            ExportTopN = 1,
            Findings =
            [
                new ReviewFinding
                {
                    Rank = 1,
                    OriginalRank = 1,
                    Product = "Windows 11",
                    IncludeInExport = true,
                    RevisedRemediation = "Patch Windows.",
                    AffectedSystems = Enumerable.Range(1, 5000)
                        .Select(i => new ReviewAffectedSystem
                        {
                            HostName = $"WORKSTATION-{i:D4}",
                            Ip = $"10.0.{i / 256}.{i % 256}",
                            VulnCount = 2
                        })
                        .ToList()
                }
            ]
        };

        try
        {
            new FlatXlsxExporter().Export(session, path);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 1000);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }
}

public class ConnectSecureDiTests
{
    [Fact]
    public void ConnectSecureClient_IsSingletonAcrossInjections()
    {
        var services = new ServiceCollection();
        services.AddVScanMagicConnectSecure();
        var provider = services.BuildServiceProvider();

        var clientA = provider.GetRequiredService<ConnectSecureClient>();
        clientA.Configure(new ConnectSecureCredentials
        {
            BaseUrl = "https://example.test",
            TenantName = "tenant",
            ClientId = "id",
            ClientSecret = "secret"
        });

        var clientB = provider.GetRequiredService<ConnectSecureClient>();
        Assert.Same(clientA, clientB);
        Assert.True(clientB.IsConfigured);
    }
}

public class AppRestartSupportTests
{
    [Theory]
    [InlineData(null, true)]
    [InlineData("127.0.0.1", true)]
    [InlineData("localhost", true)]
    [InlineData("::1", true)]
    [InlineData("0.0.0.0", false)]
    [InlineData("+", false)]
    public void IsLocalBind_AllowsLoopbackOnly(string? bind, bool expected) =>
        Assert.Equal(expected, Web.Services.AppRestartSupport.IsLocalBind(bind));

    [Fact]
    public void ResolveSrcDirectory_UsesParentOfContentRoot()
    {
        var src = Web.Services.AppRestartSupport.ResolveSrcDirectory("/tmp/vscan/src/VScanMagic.Web");
        Assert.Equal("/tmp/vscan/src", src);
    }

    [Fact]
    public void BuildLinuxDevRestartScript_EscapesSingleQuotes()
    {
        var script = Web.Services.AppRestartSupport.BuildLinuxDevRestartScript("/tmp/o'brien/src", "127.0.0.1", "8080");
        Assert.Contains("cd '/tmp/o'\\''brien/src'", script);
        Assert.Contains("export VSCANMAGIC_PORT='8080'", script);
        Assert.Contains("fuser -k 8080/tcp", script);
    }

    [Fact]
    public void BuildWindowsDevRestartScript_AvoidsInlineDoubleQuotesAndStartsDetached()
    {
        var script = Web.Services.AppRestartSupport.BuildWindowsDevRestartScript(
            @"C:\Users\dev\Documents\GitHub\vscanmagic\src",
            "127.0.0.1",
            "8080");

        Assert.DoesNotContain("\"", script);
        Assert.Contains("Start-Process -FilePath 'dotnet'", script);
        Assert.Contains("$env:VSCANMAGIC_API_BIND = '127.0.0.1'", script);
        Assert.Contains("Set-Location 'C:\\Users\\dev\\Documents\\GitHub\\vscanmagic\\src'", script);
        Assert.Contains("netstat -ano", script);
    }

    [Fact]
    public void BuildWindowsDevRestartScript_EscapesSingleQuotesInPath()
    {
        var script = Web.Services.AppRestartSupport.BuildWindowsDevRestartScript(
            @"C:\Users\o'brien\src",
            "127.0.0.1",
            "8080");

        Assert.Contains(@"Set-Location 'C:\Users\o''brien\src'", script);
    }
}

public class ReportCatalogBuilderTests
{
    [Fact]
    public void BuildGroups_GroupsByCategoryDisplay()
    {
        var groups = ReportCatalogBuilder.BuildGroups(
        [
            new StandardReportDescriptor
            {
                Id = "aaa",
                ReportType = "xlsx",
                CategoryDisplay = "All Vulnerabilities Report",
                Category = "all vulnerabilities report"
            },
            new StandardReportDescriptor
            {
                Id = "bbb",
                ReportType = "docx",
                CategoryDisplay = "All Vulnerabilities Report",
                Category = "all vulnerabilities report"
            },
            new StandardReportDescriptor
            {
                Id = "ccc",
                ReportType = "pdf",
                CategoryDisplay = "Executive Summary Report",
                Category = "executive summary report"
            }
        ]);

        Assert.Equal(2, groups.Count);
        Assert.Equal(2, groups.First(g => g.Name == "All Vulnerabilities Report").Formats.Count);
        Assert.True(groups.First(g => g.Name == "All Vulnerabilities Report").Formats.ContainsKey("xlsx"));
    }
}

public class OutlookDeliverableDraftServiceTests
{
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not-an-email")]
    public void TryValidateClientEmail_RejectsInvalid(string address)
    {
        var ok = OutlookDeliverableDraftService.TryValidateClientEmail(address, out var normalized, out var error);
        Assert.False(ok);
        Assert.Equal("", normalized);
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void TryValidateClientEmail_AcceptsValidAddress()
    {
        var ok = OutlookDeliverableDraftService.TryValidateClientEmail(
            "  client@example.com  ",
            out var normalized,
            out var error);

        Assert.True(ok);
        Assert.Equal("client@example.com", normalized);
        Assert.Equal("", error);
    }
}

public class BulkReviewJobRepositoryTests
{
    [Fact]
    public async Task SaveAndGet_RoundTripsJobPayload()
    {
        var dir = Path.Combine(Path.GetTempPath(), "VScanMagicTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        Environment.SetEnvironmentVariable("VSCANMAGIC_CONFIG_DIR", dir);
        try
        {
            using var repo = new SqliteBulkReviewJobRepository(dir);
            var job = new BulkReviewJob
            {
                ScanDate = "2026-05-21",
                Presenter = "Tester",
                Items =
                [
                    new BulkReviewJobItem { CompanyId = "1", CompanyName = "Contoso", Phase = BulkReviewItemPhase.Pending }
                ]
            };

            await repo.SaveAsync(job);
            var loaded = await repo.GetAsync(job.Id);

            Assert.NotNull(loaded);
            Assert.Equal("2026-05-21", loaded!.ScanDate);
            Assert.Single(loaded.Items);
            Assert.Equal("Contoso", loaded.Items[0].CompanyName);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            Environment.SetEnvironmentVariable("VSCANMAGIC_CONFIG_DIR", null);
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }
}

public class CheckboxRangeColumnTests
{
    [Fact]
    public void Click_ShiftClickAppliesAnchorStateToRange()
    {
        var column = new Web.Helpers.CheckboxRangeColumn();
        var values = new[] { false, false, false, false, false };

        column.Click(1, true, new MouseEventArgs { ShiftKey = false }, (i, on) => values[i] = on);
        column.Click(4, false, new MouseEventArgs { ShiftKey = true }, (i, on) => values[i] = on);

        Assert.Equal([false, true, true, true, true], values);
    }
}
