namespace VScanMagic.ConnectSecure;

using System.Globalization;

public static class AgentConnectivityHelper
{
    /// <summary>
    /// ConnectSecure Agents UI treats hosts as offline when last_ping_time is stale.
    /// One hour aligns closely with the portal lightweight/probe Online column for patch gating.
    /// </summary>
    public static readonly TimeSpan DefaultOnlineThreshold = TimeSpan.FromHours(1);

    public static bool IsOnlineFromLastPing(
        string? lastPingTime,
        DateTime? utcNow = null,
        TimeSpan? threshold = null)
    {
        if (!TryParseConnectSecureTimestamp(lastPingTime, out var ping))
            return false;

        var now = utcNow is null
            ? DateTimeOffset.UtcNow
            : new DateTimeOffset(DateTime.SpecifyKind(utcNow.Value, DateTimeKind.Utc), TimeSpan.Zero);

        return (now - ping) <= (threshold ?? DefaultOnlineThreshold);
    }

    internal static bool TryParseConnectSecureTimestamp(string? value, out DateTimeOffset timestamp)
    {
        timestamp = default;
        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (DateTimeOffset.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out timestamp))
            return true;

        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out timestamp);
    }

    public static bool IsOnlineFromAgentTimestamps(
        string? lastPingTime,
        string? lastReportedTime,
        DateTime? utcNow = null,
        TimeSpan? threshold = null)
    {
        if (!string.IsNullOrWhiteSpace(lastPingTime))
            return IsOnlineFromLastPing(lastPingTime, utcNow, threshold);

        return IsOnlineFromLastPing(lastReportedTime, utcNow, threshold);
    }

    public static string FormatAgentTypeLabel(string? agentType)
    {
        if (string.IsNullOrWhiteSpace(agentType))
            return "—";

        if (agentType.Contains("probe", StringComparison.OrdinalIgnoreCase))
            return "Probe";

        if (agentType.Contains("lightweight", StringComparison.OrdinalIgnoreCase))
            return "Lightweight";

        return agentType.Trim();
    }

    public static string FormatLastPingSummary(string? lastPingTime)
    {
        if (!TryParseConnectSecureTimestamp(lastPingTime, out var ping))
            return "no ping recorded";

        return $"last ping {ping.ToLocalTime():yyyy-MM-dd HH:mm}";
    }

    public static string FormatConnectivitySummary(bool online, string? lastPingTime) =>
        online
            ? $"Online · {FormatLastPingSummary(lastPingTime)}"
            : $"Offline · {FormatLastPingSummary(lastPingTime)}";
}
