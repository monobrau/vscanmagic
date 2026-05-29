namespace VScanMagic.ConnectSecure;

/// <summary>
/// When stable filenames are enabled, downloads overwrite canonical paths (no timestamp suffix).
/// </summary>
public sealed record ReportDownloadOptions(bool UseStableFilenames = false);
