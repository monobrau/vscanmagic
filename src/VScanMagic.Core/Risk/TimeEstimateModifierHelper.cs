namespace VScanMagic.Core.Risk;

public static class TimeEstimateModifierHelper
{
    public static bool IsTicketGenerated(bool afterHours, bool ticketGenerated, bool thirdParty) =>
        ticketGenerated || (thirdParty && afterHours);

    public static string GetModifierText(bool afterHours, bool ticketGenerated, bool thirdParty)
    {
        var isTicketGenerated = IsTicketGenerated(afterHours, ticketGenerated, thirdParty);

        if (afterHours && isTicketGenerated && thirdParty)
            return " - After-hours ticket generated for 3rd party application";
        if (afterHours && isTicketGenerated)
            return " - After-hours ticket generated";
        if (isTicketGenerated && thirdParty)
            return " - Ticket generated for 3rd party application";
        if (afterHours && thirdParty)
            return " - After-hours work required for 3rd party application, approval needed";
        if (isTicketGenerated)
            return " - Ticket generated";
        if (afterHours)
            return " - After-hours work required";
        if (thirdParty)
            return " - 3rd party application, approval needed";

        return "";
    }

    /// <summary>
    /// Modifier text for ticket/report subject lines only (no redundant "ticket generated" text).
    /// When <paramref name="afterHours"/> is true, caller prepends "After Hours - " to the subject.
    /// </summary>
    public static string GetModifierTextForSubject(bool afterHours, bool ticketGenerated, bool thirdParty)
    {
        if (IsTicketGenerated(afterHours, ticketGenerated, thirdParty))
            return "";

        if (afterHours && thirdParty)
            return " - 3rd party application";
        if (afterHours)
            return "";
        if (thirdParty)
            return " - 3rd party application";

        return "";
    }
}
