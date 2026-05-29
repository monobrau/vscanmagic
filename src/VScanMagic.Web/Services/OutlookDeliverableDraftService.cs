using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace VScanMagic.Web.Services;

/// <summary>
/// Opens a Classic Outlook compose window on the machine running VScanMagic Web,
/// with deliverable HTML above the user's default signature.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class OutlookDeliverableDraftService
{
    private const int OlMailItem = 0;
    private const int OlFormatHtml = 2;

    public bool IsSupported =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

    public static bool TryValidateClientEmail(string? address, out string normalized, out string error)
    {
        normalized = "";
        error = "";
        if (string.IsNullOrWhiteSpace(address))
        {
            error = "Enter the client email address.";
            return false;
        }

        normalized = address.Trim();
        if (!MailAddress.TryCreate(normalized, out _))
        {
            error = "Enter a valid client email address.";
            return false;
        }

        return true;
    }

    public Task<OutlookDeliverableDraftResult> OpenClientEmailDraftAsync(
        string toAddress,
        string subject,
        string htmlBodyFragment,
        CancellationToken cancellationToken = default)
    {
        if (!TryValidateClientEmail(toAddress, out var normalizedTo, out var validationError))
            return Task.FromResult(OutlookDeliverableDraftResult.Fail(validationError));

        if (string.IsNullOrWhiteSpace(subject))
            return Task.FromResult(OutlookDeliverableDraftResult.Fail("Email subject is empty."));

        if (string.IsNullOrWhiteSpace(htmlBodyFragment))
            return Task.FromResult(OutlookDeliverableDraftResult.Fail("Email body is empty."));

        if (!IsSupported)
        {
            return Task.FromResult(OutlookDeliverableDraftResult.Unavailable(
                "Open in Outlook requires Windows with Classic Outlook on the machine running VScanMagic."));
        }

        return Task.Run(() => RunSta(() =>
            OpenDraftSta(normalizedTo, subject.Trim(), htmlBodyFragment)), cancellationToken);
    }

    private static T RunSta<T>(Func<T> action)
    {
        T? result = default;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = action();
            }
            catch (Exception ex)
            {
                error = ex;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (!thread.Join(TimeSpan.FromMinutes(2)))
            return (T)(object)OutlookDeliverableDraftResult.Fail("Timed out waiting for Classic Outlook.");

        if (error is not null)
            return (T)(object)OutlookDeliverableDraftResult.Fail(error.Message);

        return result!;
    }

    private static OutlookDeliverableDraftResult OpenDraftSta(string to, string subject, string htmlBodyFragment)
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType is null)
        {
            return OutlookDeliverableDraftResult.Unavailable(
                "Classic Outlook is not installed on this machine.");
        }

        dynamic outlook = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("Could not start Classic Outlook.");

        dynamic mail = outlook.CreateItem(OlMailItem);
        mail.To = to;
        mail.Subject = subject;
        mail.BodyFormat = OlFormatHtml;

        SetBodyPreservingSignature(mail, htmlBodyFragment);
        mail.Display(false);

        return OutlookDeliverableDraftResult.Ok(
            "Opened in Classic Outlook. Click Encrypt if needed, then Send.");
    }

    private static void SetBodyPreservingSignature(dynamic mailItem, string bodyHtmlFragment)
    {
        dynamic inspector = mailItem.GetInspector();
        if (inspector is null)
            throw new InvalidOperationException("Could not get Outlook Inspector to load your signature.");

        Thread.Sleep(500);

        string signatureHtml = mailItem.HTMLBody ?? "";
        if (string.IsNullOrWhiteSpace(signatureHtml))
        {
            mailItem.Display(false);
            Thread.Sleep(700);
            signatureHtml = mailItem.HTMLBody ?? "";
        }

        var wrapped = $"""
            <div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">
            {bodyHtmlFragment}
            </div>
            <br>
            """;

        if (Regex.IsMatch(signatureHtml, @"(?is)<body[^>]*>"))
        {
            mailItem.HTMLBody = Regex.Replace(
                signatureHtml,
                @"(?is)(<body[^>]*>)",
                $"$1{wrapped}");
        }
        else
        {
            mailItem.HTMLBody = wrapped + signatureHtml;
        }
    }
}

public readonly record struct OutlookDeliverableDraftResult(
    bool Success,
    bool IsUnavailable,
    string Message)
{
    public static OutlookDeliverableDraftResult Ok(string message) => new(true, false, message);

    public static OutlookDeliverableDraftResult Fail(string message) => new(false, false, message);

    public static OutlookDeliverableDraftResult Unavailable(string message) => new(false, true, message);
}
