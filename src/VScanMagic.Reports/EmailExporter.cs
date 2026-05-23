using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class EmailExporter(TemplatesService templatesService)
{
    public (string TextPath, string EmlPath) Export(
        ReviewSession session,
        string outputDirectory,
        string companyName,
        string timestamp)
    {
        Directory.CreateDirectory(outputDirectory);
        var textPath = ReportPathResolver.GetSafeReportOutputPath(
            outputDirectory, companyName, $" Email Template_{timestamp}", "txt");
        var emlPath = ReportPathResolver.GetSafeReportOutputPath(
            outputDirectory, companyName, $" Email Template_{timestamp}", "eml");

        var templates = templatesService.Load();
        var emailContent = EmailTemplateBuilder.Build(session, templates);
        var (subject, body) = EmailTemplateBuilder.SplitSubjectAndBody(emailContent);

        File.WriteAllText(textPath, emailContent);

        var eml = new System.Text.StringBuilder();
        eml.AppendLine($"Subject: {subject}");
        eml.AppendLine("MIME-Version: 1.0");
        eml.AppendLine("Content-Type: text/plain; charset=utf-8");
        eml.AppendLine();
        eml.Append(body.Replace("\r\n", "\n").Replace("\n", Environment.NewLine));
        File.WriteAllText(emlPath, eml.ToString());

        return (textPath, emlPath);
    }
}
