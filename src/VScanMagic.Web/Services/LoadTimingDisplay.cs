using System.Diagnostics;

namespace VScanMagic.Web.Services;

public sealed class LoadTimingDisplay
{
    private readonly Stopwatch _stopwatch = new();

    public string? Label { get; private set; }
    public string? ResultText { get; private set; }

    public void Start(string label)
    {
        Label = label;
        ResultText = null;
        _stopwatch.Restart();
    }

    public void Complete(string? detail = null)
    {
        _stopwatch.Stop();
        var elapsed = _stopwatch.ElapsedMilliseconds;
        ResultText = string.IsNullOrWhiteSpace(detail)
            ? $"{Label} completed in {elapsed:N0} ms"
            : $"{Label} completed in {elapsed:N0} ms ({detail})";
    }

    public void Fail(string message)
    {
        _stopwatch.Stop();
        ResultText = $"{Label} failed after {_stopwatch.ElapsedMilliseconds:N0} ms: {message}";
    }
}
