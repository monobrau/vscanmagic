namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureOptions
{
    public int RequestsPerMinute { get; set; } = 300;
    public int RequestsPerHour { get; set; } = 2000;
}

public sealed class RateLimiter
{
    private readonly Queue<DateTimeOffset> _minuteHistory = new();
    private readonly Queue<DateTimeOffset> _hourHistory = new();
    private readonly object _lock = new();

    public async Task WaitAsync(int maxPerMinute, int maxPerHour, CancellationToken ct = default)
    {
        while (true)
        {
            ct.ThrowIfCancellationRequested();
            lock (_lock)
            {
                Prune(_minuteHistory, 60);
                Prune(_hourHistory, 3600);
                if (_minuteHistory.Count < maxPerMinute && _hourHistory.Count < maxPerHour)
                {
                    var now = DateTimeOffset.UtcNow;
                    _minuteHistory.Enqueue(now);
                    _hourHistory.Enqueue(now);
                    return;
                }
            }

            await Task.Delay(TimeSpan.FromSeconds(1), ct);
        }
    }

    private static void Prune(Queue<DateTimeOffset> history, int windowSeconds)
    {
        var cutoff = DateTimeOffset.UtcNow.AddSeconds(-windowSeconds);
        while (history.Count > 0 && history.Peek() < cutoff)
            history.Dequeue();
    }
}
