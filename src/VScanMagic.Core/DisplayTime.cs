namespace VScanMagic.Core;

/// <summary>
/// User-facing timestamps use the machine's local timezone (not UTC wall clock).
/// </summary>
public static class DisplayTime
{
    public static DateTimeOffset Now => DateTimeOffset.Now;

    public static DateTimeOffset ToLocal(DateTimeOffset value) => value.ToLocalTime();

    public static DateTimeOffset? ToLocal(DateTimeOffset? value) =>
        value is null ? null : value.Value.ToLocalTime();

    public static string Format(DateTimeOffset? value, string format = "yyyy-MM-dd HH:mm") =>
        value is null ? "" : value.Value.ToLocalTime().ToString(format);

    public static DateTimeOffset? ParseApiTimestamp(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;

        return DateTimeOffset.TryParse(
            text,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeUniversal,
            out var parsed)
            ? parsed.ToLocalTime()
            : null;
    }
}
