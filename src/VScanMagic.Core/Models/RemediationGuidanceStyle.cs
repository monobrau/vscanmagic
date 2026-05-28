namespace VScanMagic.Core.Models;

/// <summary>
/// How remediation guidance should be interpreted for a product family.
/// </summary>
public enum RemediationGuidanceStyle
{
    /// <summary>Standard patch / vendor update workflow.</summary>
    Standard = 0,

    /// <summary>
    /// Software typically self-updates (browsers, etc.). Guidance emphasizes verification
    /// and update channels; ConnectSecure fix remains the target version in OriginalFix.
    /// </summary>
    AutoUpdate = 1,

    /// <summary>Side-by-side runtime (Visual C++ redistributables).</summary>
    SideBySideRuntime = 2,

    /// <summary>Configuration or hardening change, not a version bump.</summary>
    Configuration = 3,

    /// <summary>Infrastructure or migration project (EOL OS, etc.).</summary>
    Infrastructure = 4
}
