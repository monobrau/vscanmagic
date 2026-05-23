namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureOptions
{
    public int RequestsPerMinute { get; set; } = 300;
    public int RequestsPerHour { get; set; } = 2000;
}

public sealed class RateLimiter
{
    private readonly Queue<DateTimeOffset> _history = new();
    private readonly object _lock = new();

    public async Task WaitAsync(int maxRequests, int windowSeconds, CancellationToken ct = default)
    {
        while (true)
        {
            ct.ThrowIfCancellationRequested();
            lock (_lock)
            {
                Prune(windowSeconds);
                if (_history.Count < maxRequests)
                {
                    _history.Enqueue(DateTimeOffset.UtcNow);
                    return;
                }
            }
            await Task.Delay(TimeSpan.FromSeconds(Math.Min(windowSeconds, 5)), ct);
        }
    }

    private void Prune(int windowSeconds)
    {
        var cutoff = DateTimeOffset.UtcNow.AddSeconds(-windowSeconds);
        while (_history.Count > 0 && _history.Peek() < cutoff)
            _history.Dequeue();
    }
}
