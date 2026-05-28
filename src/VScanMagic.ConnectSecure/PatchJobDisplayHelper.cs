using VScanMagic.Core.Services;

namespace VScanMagic.ConnectSecure;

public static class PatchJobDisplayHelper
{
    public static string? ResolveProduct(
        PatchActivityEntry? entry,
        PatchJobCorrelationHelper.ParsedConnectSecureJob? remote)
    {
        foreach (var candidate in new[]
                 {
                     entry?.Product,
                     remote?.ProductName,
                     TryExtractProductFromDescription(entry?.Description),
                     TryExtractProductFromDescription(remote?.Description)
                 })
        {
            if (string.IsNullOrWhiteSpace(candidate) || IsGenericPatchLabel(candidate))
                continue;

            return candidate.Trim();
        }

        return null;
    }

    public static string ResolveApplication(PatchJobEntry job)
    {
        if (!string.IsNullOrWhiteSpace(job.Product))
            return job.Product;

        var fromDescription = TryExtractProductFromDescription(job.Description);
        if (!string.IsNullOrWhiteSpace(fromDescription))
            return fromDescription;

        if (job.Type.Contains("os", StringComparison.OrdinalIgnoreCase))
            return "OS update";

        if (!string.IsNullOrWhiteSpace(job.Description) && !IsGenericPatchLabel(job.Description))
            return job.Description;

        return "—";
    }

    public static string ResolveDetail(PatchJobEntry job)
    {
        var application = job.Product ?? TryExtractProductFromDescription(job.Description);
        var detail = StripLeadingProduct(job.Description, application);

        if (!string.IsNullOrWhiteSpace(detail))
            return detail;

        return string.IsNullOrWhiteSpace(job.HostName) ? "—" : job.HostName;
    }

    public static string BuildMergedDescription(
        PatchActivityEntry entry,
        PatchJobCorrelationHelper.ParsedConnectSecureJob? remote,
        string? product)
    {
        var local = entry.Description?.Trim();
        var remoteDesc = remote?.Description?.Trim();

        if (!string.IsNullOrWhiteSpace(local) &&
            local.Contains(" on ", StringComparison.OrdinalIgnoreCase))
            return local;

        if (!string.IsNullOrWhiteSpace(remoteDesc) &&
            remoteDesc.Contains("Success:", StringComparison.OrdinalIgnoreCase))
            return remoteDesc;

        if (!string.IsNullOrWhiteSpace(local) && !IsGenericPatchLabel(local))
            return local;

        if (!string.IsNullOrWhiteSpace(remoteDesc))
            return remoteDesc;

        if (!string.IsNullOrWhiteSpace(product) && !string.IsNullOrWhiteSpace(entry.HostName))
            return $"{product} on {entry.HostName}";

        return entry.HostName ?? remote?.HostName ?? remoteDesc ?? local ?? "";
    }

    internal static string? TryExtractProductFromDescription(string? description)
    {
        if (string.IsNullOrWhiteSpace(description))
            return null;

        var text = description.Trim();
        if (IsGenericPatchLabel(text))
            return null;

        var onIndex = text.IndexOf(" on ", StringComparison.OrdinalIgnoreCase);
        if (onIndex > 0)
            return text[..onIndex].Trim();

        var dashIndex = text.IndexOf(" — ", StringComparison.Ordinal);
        if (dashIndex > 0)
            return text[..dashIndex].Trim();

        return null;
    }

    private static string StripLeadingProduct(string? description, string? product)
    {
        if (string.IsNullOrWhiteSpace(description))
            return "";

        var detail = description.Trim();
        if (string.IsNullOrWhiteSpace(product))
            return detail;

        if (!detail.StartsWith(product, StringComparison.OrdinalIgnoreCase))
            return detail;

        detail = detail[product.Length..].TrimStart();
        if (detail.StartsWith("—", StringComparison.Ordinal))
            detail = detail[1..].TrimStart();
        if (detail.StartsWith("-", StringComparison.Ordinal))
            detail = detail[1..].TrimStart();
        if (detail.StartsWith("on ", StringComparison.OrdinalIgnoreCase))
            detail = detail[3..].TrimStart();

        return detail;
    }

    private static bool IsGenericPatchLabel(string value) =>
        value.Equals("Patch request", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("Application Patch", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("Scheduled Application Patch", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("OS Patch", StringComparison.OrdinalIgnoreCase);
}
