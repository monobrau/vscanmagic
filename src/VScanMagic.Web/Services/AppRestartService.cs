using System.Diagnostics;

namespace VScanMagic.Web.Services;

public sealed class AppRestartService(
    IHostApplicationLifetime lifetime,
    IWebHostEnvironment environment,
    ILogger<AppRestartService> logger)
{
    public static bool IsRestartPermitted() =>
        AppRestartSupport.IsLocalBind(Environment.GetEnvironmentVariable("VSCANMAGIC_API_BIND"));

    public Task ScheduleRestartAsync(CancellationToken ct = default)
    {
        if (!IsRestartPermitted())
            throw new InvalidOperationException("Restart is only allowed when VSCANMAGIC_API_BIND is loopback (127.0.0.1, localhost, or ::1).");

        var bind = Environment.GetEnvironmentVariable("VSCANMAGIC_API_BIND") ?? "127.0.0.1";
        var port = Environment.GetEnvironmentVariable("VSCANMAGIC_PORT") ?? "8080";

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(750, ct).ConfigureAwait(false);
                StartReplacementProcess(bind, port);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to start replacement VScanMagic process");
            }
            finally
            {
                lifetime.StopApplication();
            }
        }, CancellationToken.None);

        return Task.CompletedTask;
    }

    private void StartReplacementProcess(string bind, string port)
    {
        var contentRoot = environment.ContentRootPath;
        var projectFile = Path.Combine(contentRoot, "VScanMagic.Web.csproj");
        if (File.Exists(projectFile))
        {
            var srcDir = AppRestartSupport.ResolveSrcDirectory(contentRoot);
            StartDetachedShell(AppRestartSupport.BuildLinuxDevRestartScript(srcDir, bind, port),
                AppRestartSupport.BuildWindowsDevRestartScript(srcDir, bind, port));
            logger.LogInformation("Scheduled dev restart from {SrcDir} on http://{Bind}:{Port}", srcDir, bind, port);
            return;
        }

        var exePath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Cannot restart: process path is unknown (publish the app or run via dotnet run).");

        var psi = new ProcessStartInfo
        {
            FileName = exePath,
            WorkingDirectory = contentRoot,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        psi.Environment["VSCANMAGIC_API_BIND"] = bind;
        psi.Environment["VSCANMAGIC_PORT"] = port;

        Process.Start(psi);
        logger.LogInformation("Scheduled published restart of {Exe} on http://{Bind}:{Port}", exePath, bind, port);
    }

    private static void StartDetachedShell(string linuxScript, string windowsScript)
    {
        if (OperatingSystem.IsWindows())
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command \"{windowsScript}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
            });
            return;
        }

        var escaped = AppRestartSupport.EscapeSingleQuotedShell(linuxScript);
        Process.Start(new ProcessStartInfo
        {
            FileName = "/bin/bash",
            Arguments = $"-c \"nohup bash -c '{escaped}' </dev/null >/dev/null 2>&1 &\"",
            UseShellExecute = false,
            CreateNoWindow = true,
        });
    }
}
