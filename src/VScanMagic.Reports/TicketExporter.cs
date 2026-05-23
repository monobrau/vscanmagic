using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class TicketExporter(RemediationRuleService remediationRules)
{
    public string Export(ReviewSession session, string? reportsPathPartial = null) =>
        TicketInstructionBuilder.BuildPlainTextDocument(session, remediationRules, reportsPathPartial);

    public void ExportToFile(ReviewSession session, string outputPath, string? reportsPathPartial = null)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(outputPath, Export(session, reportsPathPartial));
    }
}
