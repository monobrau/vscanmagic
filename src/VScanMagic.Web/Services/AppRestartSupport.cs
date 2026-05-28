namespace VScanMagic.Web.Services;

public static class AppRestartSupport
{
    public static bool IsLocalBind(string? bind)
    {
        var value = (bind ?? "127.0.0.1").Trim();
        return value.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase)
            || value.Equals("localhost", StringComparison.OrdinalIgnoreCase)
            || value.Equals("::1", StringComparison.OrdinalIgnoreCase);
    }

    public static string ResolveSrcDirectory(string contentRootPath)
    {
        var webDir = Path.GetFullPath(contentRootPath);
        var srcDir = Directory.GetParent(webDir)?.FullName;
        if (string.IsNullOrWhiteSpace(srcDir))
            throw new InvalidOperationException($"Could not resolve src directory from content root '{webDir}'.");

        return srcDir;
    }

    public static string BuildLinuxDevRestartScript(string srcDirectory, string bind, string port)
    {
        var escapedSrc = EscapeSingleQuotedShell(srcDirectory);
        var escapedBind = EscapeSingleQuotedShell(bind);
        var escapedPort = EscapeSingleQuotedShell(port);
        return $"""
            export PATH="$HOME/.dotnet:$PATH"
            export VSCANMAGIC_API_BIND='{escapedBind}'
            export VSCANMAGIC_PORT='{escapedPort}'
            sleep 2
            fuser -k {port}/tcp 2>/dev/null || true
            cd '{escapedSrc}'
            exec dotnet run --project VScanMagic.Web -c Release
            """;
    }

    public static string BuildWindowsDevRestartScript(string srcDirectory, string bind, string port)
    {
        if (!int.TryParse(port, out var portNumber) || portNumber is <= 0 or > 65535)
            throw new ArgumentOutOfRangeException(nameof(port), port, "Port must be a valid TCP port number.");

        var escapedSrc = EscapePowerShellSingleQuoted(srcDirectory);
        var escapedBind = EscapePowerShellSingleQuoted(bind);

        return $$"""
            $ErrorActionPreference = 'SilentlyContinue'
            $dotnetUser = Join-Path $env:USERPROFILE '.dotnet'
            $env:PATH = [string]::Join(';', $dotnetUser, $env:PATH)
            $env:VSCANMAGIC_API_BIND = '{{escapedBind}}'
            $env:VSCANMAGIC_PORT = '{{portNumber}}'
            Start-Sleep -Seconds 2
            $port = {{portNumber}}
            Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
            netstat -ano | Select-String (':' + $port + ' ') | ForEach-Object {
                if ($_ -match 'LISTENING\s+(\d+)\s*$') {
                    Stop-Process -Id ([int]$Matches[1]) -Force -ErrorAction SilentlyContinue
                }
            }
            Set-Location '{{escapedSrc}}'
            Start-Process -FilePath 'dotnet' -ArgumentList @('run','--project','VScanMagic.Web','-c','Release') -WorkingDirectory '{{escapedSrc}}' -WindowStyle Hidden
            """;
    }

    public static string EscapeSingleQuotedShell(string value) =>
        value.Replace("'", "'\\''");

    public static string EscapePowerShellSingleQuoted(string value) =>
        value.Replace("'", "''");
}
