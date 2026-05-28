using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VScanMagic.Web.Services;

/// <summary>
/// Opens a native folder picker on the machine running the web app (local Blazor Server use).
/// </summary>
public sealed class NativeFolderPickerService
{
    public async Task<FolderPickerResult> PickFolderAsync(string? initialPath, CancellationToken cancellationToken = default)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return await PickFolderWindowsAsync(initialPath, cancellationToken).ConfigureAwait(false);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            return await PickFolderMacOsAsync(initialPath, cancellationToken).ConfigureAwait(false);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            return await PickFolderLinuxAsync(initialPath, cancellationToken).ConfigureAwait(false);

        return FolderPickerResult.Unavailable("Folder browse is not supported on this operating system.");
    }

    private static async Task<FolderPickerResult> PickFolderWindowsAsync(string? initialPath, CancellationToken cancellationToken)
    {
        var start = string.IsNullOrWhiteSpace(initialPath) ? "" : Path.GetFullPath(initialPath.Trim());
        var script = new StringBuilder();
        script.AppendLine("Add-Type -AssemblyName System.Windows.Forms");
        script.AppendLine("$dialog = New-Object System.Windows.Forms.FolderBrowserDialog");
        script.AppendLine("$dialog.Description = 'Select output directory'");
        script.AppendLine("$dialog.ShowNewFolderButton = $true");
        if (!string.IsNullOrEmpty(start) && Directory.Exists(start))
            script.AppendLine($"$dialog.SelectedPath = '{EscapePowerShellSingleQuoted(start)}'");
        script.AppendLine("if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {");
        script.AppendLine("  [Console]::Out.Write($dialog.SelectedPath)");
        script.AppendLine("}");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-STA -NoProfile -ExecutionPolicy Bypass -Command \"{EscapeProcessArgument(script.ToString())}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process is null)
            return FolderPickerResult.Unavailable("Could not start PowerShell for folder browse.");

        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        var output = (await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false)).Trim();
        if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            return FolderPickerResult.UserCancelled();

        return FolderPickerResult.Ok(output);
    }

    private static async Task<FolderPickerResult> PickFolderLinuxAsync(string? initialPath, CancellationToken cancellationToken)
    {
        var start = ResolveInitialDirectory(initialPath);
        if (File.Exists("/usr/bin/zenity"))
            return await RunPickerProcessAsync("/usr/bin/zenity",
                $"--file-selection --directory --title=Select output directory --filename={QuoteShell(start)}/",
                cancellationToken).ConfigureAwait(false);

        if (File.Exists("/usr/bin/kdialog"))
            return await RunPickerProcessAsync("/usr/bin/kdialog",
                $"--getexistingdirectory {QuoteShell(start)} --title \"Select output directory\"",
                cancellationToken).ConfigureAwait(false);

        return FolderPickerResult.Unavailable(
            "Install zenity or kdialog for folder browse on Linux, or type the path manually.");
    }

    private static async Task<FolderPickerResult> PickFolderMacOsAsync(string? initialPath, CancellationToken cancellationToken)
    {
        var start = ResolveInitialDirectory(initialPath);
        var escaped = EscapeAppleScript(start);
        var script =
            $"POSIX path of (choose folder with prompt \"Select output directory\" default location POSIX file \"{escaped}\")";
        var psi = new ProcessStartInfo
        {
            FileName = "/usr/bin/osascript",
            Arguments = $"-e {QuoteShell(script)}",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        return await RunPickerProcessAsync(psi, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<FolderPickerResult> RunPickerProcessAsync(string fileName, string arguments, CancellationToken cancellationToken)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        return await RunPickerProcessAsync(psi, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<FolderPickerResult> RunPickerProcessAsync(ProcessStartInfo psi, CancellationToken cancellationToken)
    {
        using var process = Process.Start(psi);
        if (process is null)
            return FolderPickerResult.Unavailable("Could not start folder picker.");

        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        var output = (await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false)).Trim();
        if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            return FolderPickerResult.UserCancelled();

        return FolderPickerResult.Ok(output.TrimEnd('/'));
    }

    private static string ResolveInitialDirectory(string? initialPath)
    {
        if (!string.IsNullOrWhiteSpace(initialPath))
        {
            var trimmed = initialPath.Trim();
            if (Directory.Exists(trimmed))
                return Path.GetFullPath(trimmed);
            var parent = Path.GetDirectoryName(trimmed);
            if (!string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent))
                return Path.GetFullPath(parent);
        }

        return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }

    private static string EscapePowerShellSingleQuoted(string value) =>
        value.Replace("'", "''", StringComparison.Ordinal);

    private static string EscapeProcessArgument(string value) =>
        value.Replace("\"", "\\\"");

    private static string EscapeAppleScript(string path) =>
        path.Replace("\\", "\\\\").Replace("\"", "\\\"");

    private static string QuoteShell(string value) =>
        "'" + value.Replace("'", "'\"'\"'") + "'";
}

public readonly record struct FolderPickerResult(bool Success, bool WasCancelled, string? Path, string? Error)
{
    public static FolderPickerResult Ok(string path) => new(true, false, path, null);

    public static FolderPickerResult UserCancelled() => new(false, true, null, null);

    public static FolderPickerResult Unavailable(string message) => new(false, false, null, message);
}
