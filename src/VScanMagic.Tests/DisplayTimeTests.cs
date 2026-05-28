using VScanMagic.Core;

namespace VScanMagic.Tests;

public sealed class DisplayTimeTests
{
    [Fact]
    public void ParseApiTimestamp_AssumesUtcAndConvertsToLocal()
    {
        var parsed = DisplayTime.ParseApiTimestamp("2026-05-25T18:30:00Z");
        Assert.NotNull(parsed);
        Assert.Equal(TimeZoneInfo.Local.GetUtcOffset(parsed.Value), parsed.Value.Offset);
    }

    [Fact]
    public void Format_UsesLocalWallClock()
    {
        var utc = new DateTimeOffset(2026, 5, 25, 18, 30, 0, TimeSpan.Zero);
        var formatted = DisplayTime.Format(utc);
        Assert.Equal(utc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"), formatted);
    }
}
