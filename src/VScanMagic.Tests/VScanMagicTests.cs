using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using System.Text.Json;
using Microsoft.AspNetCore.Components.Web;
using VScanMagic.ConnectSecure;
using VScanMagic.Core.Configuration;
using VScanMagic.Core.Models;
using VScanMagic.Data.Parsing;
using VScanMagic.Data.Scoring;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Models;
using VScanMagic.Reports;

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
            new Reports.DocxReviewExporter().Export(session, path);
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

        var html = new CombinedReportHtmlExporter(new TemplatesService(), new RemediationRuleService()).BuildHtml(session);

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
        Assert.Contains("Updates Required", subject);
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
        var wordGuidance = rules.GetGuidance("Microsoft Edge", forWord: true);
        var finding = new ReviewFinding
        {
            Product = "Microsoft Edge",
            OriginalRemediation = wordGuidance,
            RevisedRemediation = wordGuidance
        };

        var body = TicketInstructionBuilder.BuildBodyText(finding, rules);

        Assert.Contains("Remediation Instructions:", body);
        Assert.Contains("- Determine device/software identity", body);
        Assert.DoesNotContain("If available via ConnectWise Automate/RMM or scripting, deploy updates using the patch management system", body);
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

        Assert.Contains("Consolidated .NET finding", text);
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

        Assert.Contains("- Ticket created for Adobe Reader", notes);
    }
    [Fact]
    public void IsCveOnlyProduct_DetectsBareCveIds()
    {
        Assert.True(CveReferenceHelper.IsCveOnlyProduct("CVE-2021-33558"));
        Assert.False(CveReferenceHelper.IsCveOnlyProduct("Windows 11"));
        Assert.False(CveReferenceHelper.IsCveOnlyProduct("Google Chrome CVE-2021-33558"));
    }

    [Fact]
    public void GetTopVulnerabilities_ExcludesCveOnlyProducts()
    {
        var options = new VScanMagicOptions();
        var scorer = new TopVulnerabilityScorer(options);
        var data = new List<VulnerabilityRecord>
        {
            new() { Product = "CVE-2021-33558", HostName = "PC1", Critical = 1, VulnerabilityCount = 1, EpssScore = 0.9 },
            new() { Product = "Google Chrome", HostName = "PC1", High = 3, VulnerabilityCount = 3, EpssScore = 0.5 }
        };

        var top = scorer.GetTopVulnerabilities(data, new ReportFilters { TopN = 10, IncludeCritical = true, IncludeHigh = true });
        Assert.Single(top);
        Assert.Equal("Google Chrome", top[0].Product);
    }

    [Fact]
    public void EnsureLegacyFields_RemovesCveOnlyFindings()
    {
        var session = new ReviewSession
        {
            ExportTopN = 10,
            Findings =
            [
                new ReviewFinding { OriginalRank = 1, Rank = 1, Product = "Windows 11", IncludeInExport = true, CveIds = "CVE-2024-1" },
                new ReviewFinding { OriginalRank = 2, Rank = 2, Product = "CVE-2021-33558", IncludeInExport = false, CveIds = "CVE-2021-33558" }
            ]
        };

        ReviewSessionRanker.EnsureLegacyFields(session);

        Assert.Single(session.Findings);
        Assert.Equal("Windows 11", session.Findings[0].Product);
        Assert.Equal("", session.Findings[0].CveIds);
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
    [InlineData("Microsoft Edge", true, "Updates Required")]
    [InlineData("Microsoft Edge", false, "Update Required")]
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
        var wrapper = "/home/cknospe/Documents/VScanMagic/Downloads/Accurate Metal Products Inc - All Vulnerabilities Report_2026-05-23_15-53-06.xlsx";
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
