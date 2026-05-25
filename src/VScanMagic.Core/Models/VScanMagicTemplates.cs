namespace VScanMagic.Core.Models;

public sealed class VScanMagicTemplates
{
    /// <summary>RMIT / CMIT (hourly) clients — approval required before remediation.</summary>
    public EmailTemplateSettings EmailTemplate { get; set; } = EmailTemplateSettings.CreateRmitDefault();

    /// <summary>RMIT+ agreement clients — tickets generated for covered items.</summary>
    public EmailTemplateSettings EmailTemplateRmitPlus { get; set; } = EmailTemplateSettings.CreateRmitPlusDefault();

    public TicketNotesTemplateSettings TicketNotes { get; set; } = TicketNotesTemplateSettings.CreateDefault();

    public EmailTemplateSettings ResolveEmailTemplate(bool isRmitPlus) =>
        isRmitPlus ? EmailTemplateRmitPlus : EmailTemplate;
}

public sealed class EmailTemplateSettings
{
    public string SubjectFormat { get; set; } = "{Year} Q{Quarter} Vulnerability Scan Follow Up";
    public string Body { get; set; } = DefaultBody;

    public static EmailTemplateSettings CreateRmitDefault() => new();

    public static EmailTemplateSettings CreateRmitPlusDefault() => new()
    {
        SubjectFormat = "{Year} Q{Quarter} Vulnerability Scan Follow Up",
        Body = DefaultBody
    };

    private const string DefaultBody = """
        Subject: {Year} Q{Quarter} Vulnerability Scan Follow Up

        Good {Greeting},

        Your quarterly vulnerability scan report has been completed and is available in your client folder.

        Recommended remediation priorities ({TopNLabel}):
        {TopNReportLink}

        Complete report package:
        {ReportsFolderLink}

        The folder contains the following reports:
        • Pending Remediation EPSS Score Report – Classifies vulnerabilities by Exploit Prediction Scoring System (EPSS), which measures the likelihood of exploitation within 30 days (scale 0–1.0, with 1.0 being most critical).
        • All Vulnerabilities Report – A comprehensive list of all detected vulnerabilities (internal and external), from critical to low severity.
        • Executive Summary Report – A high-level overview of your security posture and network information.
        • External Scan – Detected vulnerabilities and services exposed to the internet.
        • Suppressed Vulnerabilities Report – Vulnerabilities that have been suppressed (e.g., false positives or accepted risk) and will not appear on future remediation lists.

        Not all vulnerabilities may be feasible to remediate depending on business or technical constraints.

        Schedule time with me
        {SchedulingLink}

        {NoteText}

        We appreciate your commitment to security. Addressing these vulnerabilities is essential for maintaining the protection of your systems.

        Sincerely,

        {PreparedBy}
        """;
}

public sealed class TicketNotesTemplateSettings
{
    public string StepsBeforeTickets { get; set; } = DefaultStepsBefore;
    public string StepsAfterTickets { get; set; } = DefaultStepsAfter;
    public string ResolvedQuestion { get; set; } = "Is the task resolved?";
    public string ResolvedAnswer { get; set; } = "Yes - completed";
    public string NextStepsQuestion { get; set; } = "Next step(s)";
    public string NextStepsText { get; set; } = "TimeZest meeting request has been sent. Please select a time to meet if you would like to discuss this further.";

    public static TicketNotesTemplateSettings CreateDefault() => new();

    private const string DefaultStepsBefore = """
        - Examined lightweight agents
        - Verified probe setup
        - Checked agent/probe count compared to other systems
        - Examined credential mappings
        - Examined external assets
        - Checked nmap interface on probe
        - Verified deprecated item list
        - Created all reports
        - Assessed reports
        - {ReportStepLine}
        """;

    private const string DefaultStepsAfter = """
        - Sent secure email with reports to contact
        - Sent TimeZest meeting request
        """;
}
