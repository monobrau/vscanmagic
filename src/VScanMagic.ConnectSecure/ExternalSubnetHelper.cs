using System.Net;
using System.Text.RegularExpressions;

namespace VScanMagic.ConnectSecure;

public static partial class ExternalSubnetHelper
{
    [GeneratedRegex(@"^(\d+\.\d+\.\d+\.\d+)/(\d+)$")]
    private static partial Regex CidrRegex();

    [GeneratedRegex(@"^(\d+\.\d+\.\d+\.\d+)/(\d+\.\d+\.\d+\.\d+)$")]
    private static partial Regex CidrWithDottedMaskRegex();

    [GeneratedRegex(@"^\d+\.\d+\.\d+\.\d+/$")]
    private static partial Regex IncompleteCidrRegex();

    [GeneratedRegex(@"^\d+\.\d+\.\d+\.\d+$")]
    private static partial Regex IpRegex();

    private static readonly Regex DottedMaskRegex = new(
        @"^\d{1,3}(\.\d{1,3}){3}$",
        RegexOptions.Compiled);

    public sealed record SubnetBounds(
        string Network,
        string Gateway,
        string FirstUsable,
        string LastUsable,
        string Broadcast);

    public sealed record ExternalScanTargetValidationResult(
        bool IsValid,
        IReadOnlyList<string> Errors,
        IReadOnlyList<string> ScanIps,
        string Address,
        string TargetIp);

    public static IReadOnlyList<string> ValidateExternalTargets(IEnumerable<string> targets, string cidr)
    {
        var bounds = GetSubnetBounds(cidr);
        if (bounds is null)
            return [];

        var issues = new List<string>();
        foreach (var target in targets)
        {
            var t = target.Replace(" ", "").Trim();
            if (string.IsNullOrWhiteSpace(t))
                continue;
            if (string.Equals(t, bounds.Network, StringComparison.Ordinal))
                issues.Add($"Target includes network address: {t}");
            if (string.Equals(t, bounds.Gateway, StringComparison.Ordinal))
                issues.Add($"Target includes ISP default gateway: {t}");
            if (string.Equals(t, bounds.Broadcast, StringComparison.Ordinal))
                issues.Add($"Target includes broadcast address: {t}");
        }

        return issues;
    }

    public static ExternalScanTargetValidationResult ParseAndValidateScanInput(string input)
    {
        input = NormalizeScanInput(input);
        var segments = SplitSegments(input);
        if (segments.Count == 0)
        {
            return new ExternalScanTargetValidationResult(
                false,
                ["Enter a CIDR subnet (e.g. 10.0.0.0/24) or IP with subnet mask (e.g. 10.0.0.0 255.255.255.0)."],
                [],
                "",
                "");
        }

        var tokens = new List<string>();
        var errors = new List<string>();

        foreach (var segment in segments)
        {
            foreach (var expanded in ExpandSegment(segment))
            {
                errors.AddRange(expanded.Errors);
                if (!string.IsNullOrWhiteSpace(expanded.Token))
                    tokens.Add(expanded.Token);
            }
        }

        if (tokens.Count == 0 && errors.Count == 0)
            errors.Add("No scannable targets were resolved from the input.");

        var cidrTokens = tokens.Where(IsCidr).ToList();
        var ipTokens = tokens.Where(t => IsIp(t) && !IsCidr(t)).ToList();

        var scanIps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var cidr in cidrTokens)
        {
            var expanded = ExpandCidrToUsableIps(cidr);
            if (expanded.Count == 0)
                errors.Add($"Subnet {cidr} has no scannable usable addresses after excluding network, gateway, and broadcast.");
            foreach (var ip in expanded)
                scanIps.Add(ip);
        }

        foreach (var ip in ipTokens)
        {
            var reservedBy = FindReservedCidrMatch(ip, cidrTokens);
            if (reservedBy is not null)
            {
                var bounds = GetSubnetBounds(reservedBy)!;
                if (string.Equals(ip, bounds.Network, StringComparison.OrdinalIgnoreCase))
                    errors.Add($"Target {ip} is the network address for {reservedBy}.");
                else if (string.Equals(ip, bounds.Gateway, StringComparison.OrdinalIgnoreCase))
                    errors.Add($"Target {ip} is the ISP default gateway for {reservedBy}.");
                else if (string.Equals(ip, bounds.Broadcast, StringComparison.OrdinalIgnoreCase))
                    errors.Add($"Target {ip} is the broadcast address for {reservedBy}.");
                else if (!scanIps.Contains(ip))
                    scanIps.Add(ip);
            }
            else
            {
                scanIps.Add(ip);
            }
        }

        if (scanIps.Count == 0 && errors.Count == 0)
            errors.Add("No scannable IP addresses were resolved from the input.");

        var ordered = scanIps.OrderBy(IpSortKey).Select(IpSortKeyToString).ToList();
        var address = cidrTokens.Count == 1 && ipTokens.Count == 0
            ? cidrTokens[0]
            : string.Join(", ", tokens);

        return new ExternalScanTargetValidationResult(
            errors.Count == 0 && ordered.Count > 0,
            errors,
            ordered,
            address,
            string.Join(", ", ordered));
    }

    public static IReadOnlyList<string> ExpandCidrToUsableIps(string cidr)
    {
        var bounds = GetSubnetBounds(cidr);
        if (bounds is null)
            return [];

        if (!TryParseIpToUInt(bounds.FirstUsable, out var first) ||
            !TryParseIpToUInt(bounds.LastUsable, out var last) ||
            first > last)
            return [];

        var results = new List<string>();
        for (var value = first; value <= last; value++)
            results.Add(UIntToIp(value));
        return results;
    }

    public static SubnetBounds? GetSubnetBounds(string cidr)
    {
        var match = CidrRegex().Match(cidr.Trim());
        if (!match.Success)
            return null;

        if (!IPAddress.TryParse(match.Groups[1].Value, out var ip) ||
            !int.TryParse(match.Groups[2].Value, out var prefix) ||
            prefix is < 0 or > 32)
            return null;

        var networkLong = GetNetworkLong(ip, prefix);
        var broadcastLong = GetBroadcastLong(networkLong, prefix);
        var hostCount = broadcastLong - networkLong + 1;

        if (hostCount == 1)
        {
            var single = UIntToIp(networkLong);
            return new SubnetBounds(single, single, single, single, single);
        }

        if (hostCount == 2)
        {
            var first = UIntToIp(networkLong);
            var second = UIntToIp(broadcastLong);
            return new SubnetBounds(first, first, first, second, second);
        }

        var network = UIntToIp(networkLong);
        var gateway = UIntToIp(networkLong + 1);
        var firstUsable = UIntToIp(networkLong + 2);
        var lastUsable = UIntToIp(broadcastLong - 1);
        var broadcast = UIntToIp(broadcastLong);
        return new SubnetBounds(network, gateway, firstUsable, lastUsable, broadcast);
    }

    public static string DescribeExpandedRange(string input)
    {
        var validation = ParseAndValidateScanInput(input);
        if (!validation.IsValid || validation.ScanIps.Count == 0)
            return "";

        if (validation.ScanIps.Count == 1)
            return $"Scans 1 address: {validation.ScanIps[0]}";

        return $"Scans {validation.ScanIps.Count} addresses: {validation.ScanIps[0]} – {validation.ScanIps[^1]}";
    }

    public static bool IsIncompleteScanInput(string input)
    {
        var trimmed = NormalizeScanInput(input);
        if (string.IsNullOrWhiteSpace(trimmed))
            return false;

        if (IncompleteCidrRegex().IsMatch(trimmed))
            return true;

        if (trimmed.EndsWith('|'))
            return true;

        var parts = trimmed
            .Replace('|', ' ')
            .Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 2 && IpRegex().IsMatch(parts[0]) && parts[1].Contains('.'))
        {
            var octets = parts[1].Split('.', StringSplitOptions.RemoveEmptyEntries);
            if (octets.Length is > 0 and < 4)
                return true;
        }

        return false;
    }

    public static string NormalizeScanInput(string input) =>
        input.Trim()
            .Replace('\u2044', '/')
            .Replace('\u2215', '/');

    public static bool LooksLikeSubnetInput(string input)
    {
        input = NormalizeScanInput(input);
        if (string.IsNullOrWhiteSpace(input))
            return false;

        if (input.Contains('/'))
            return true;

        var parts = input
            .Replace('|', ' ')
            .Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return parts.Length >= 2 && IpRegex().IsMatch(parts[0]);
    }

    public static IReadOnlyList<string> ValidateScanInputForUi(string input, bool strict = true)
    {
        input = NormalizeScanInput(input);
        if (!strict && IsIncompleteScanInput(input))
            return [];

        var validation = ParseAndValidateScanInput(input);
        if (!validation.IsValid)
            return validation.Errors.ToList();

        if (LooksLikeSubnetInput(input) && validation.ScanIps.Count <= 1)
        {
            return
            [
                "Could not parse a subnet from that input. Use CIDR (192.168.50.0/24) or IP + mask (192.168.50.0 255.255.255.0)."
            ];
        }

        return [];
    }

    private static string? FindReservedCidrMatch(string ip, IReadOnlyList<string> cidrs)
    {
        foreach (var cidr in cidrs)
        {
            if (IpInCidr(ip, cidr))
                return cidr;
        }

        return null;
    }

    private static bool IpInCidr(string ip, string cidr)
    {
        var match = CidrRegex().Match(cidr.Trim());
        if (!match.Success ||
            !IPAddress.TryParse(ip, out var ipAddress) ||
            !IPAddress.TryParse(match.Groups[1].Value, out var networkIp) ||
            !int.TryParse(match.Groups[2].Value, out var prefix))
            return false;

        var networkLong = GetNetworkLong(networkIp, prefix);
        var broadcastLong = GetBroadcastLong(networkLong, prefix);
        var value = IpToUInt(ipAddress);
        return value >= networkLong && value <= broadcastLong;
    }

    private static List<string> SplitSegments(string input) =>
        input
            .Split([',', ';', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(segment => !string.IsNullOrWhiteSpace(segment))
            .ToList();

    private static IEnumerable<(string? Token, List<string> Errors)> ExpandSegment(string segment)
    {
        if (TryNormalizeSegment(segment, out var normalized, out var errors) &&
            !string.IsNullOrWhiteSpace(normalized))
        {
            yield return (normalized, errors);
            yield break;
        }

        if (errors.Count > 0)
        {
            yield return (null, errors);
            yield break;
        }

        var parts = segment
            .Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length <= 1)
        {
            yield return (null, [$"Invalid target: {segment.Trim()}"]);
            yield break;
        }

        foreach (var part in parts)
        {
            if (TryNormalizeSegment(part, out var token, out var partErrors) &&
                !string.IsNullOrWhiteSpace(token))
                yield return (token, partErrors);
            else
                yield return (null, partErrors.Count > 0 ? partErrors : [$"Invalid target: {part.Trim()}"]);
        }
    }

    internal static bool TryNormalizeSegment(string segment, out string normalized, out List<string> errors)
    {
        normalized = "";
        errors = [];
        segment = segment.Trim();
        if (string.IsNullOrWhiteSpace(segment))
            return false;

        if (IsCidr(segment))
        {
            normalized = segment;
            return true;
        }

        if (IncompleteCidrRegex().IsMatch(segment))
        {
            errors.Add("Enter a prefix length after / (e.g. 192.168.54.0/24).");
            return false;
        }

        var dottedSlash = CidrWithDottedMaskRegex().Match(segment);
        if (dottedSlash.Success && IPAddress.TryParse(dottedSlash.Groups[1].Value, out _))
        {
            if (TryMaskToPrefix(dottedSlash.Groups[2].Value, out var dottedPrefix, out var dottedError))
            {
                normalized = $"{dottedSlash.Groups[1].Value}/{dottedPrefix}";
                return true;
            }

            if (dottedError is not null)
            {
                errors.Add(dottedError);
                return false;
            }
        }

        var slashIndex = segment.IndexOf('/');
        if (slashIndex > 0)
        {
            var ipPart = segment[..slashIndex].Trim();
            var maskPart = segment[(slashIndex + 1)..].Trim();
            if (IpRegex().IsMatch(ipPart))
            {
                if (TryMaskToPrefix(maskPart, out var slashPrefix, out var slashError))
                {
                    normalized = $"{ipPart}/{slashPrefix}";
                    return true;
                }

                if (slashError is not null)
                {
                    errors.Add(slashError);
                    return false;
                }
            }
        }

        var cleaned = segment.Replace('|', ' ').Trim();
        var pairParts = cleaned
            .Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (pairParts.Length == 2 && IpRegex().IsMatch(pairParts[0]))
        {
            if (TryMaskToPrefix(pairParts[1], out var pairPrefix, out var pairError))
            {
                normalized = $"{pairParts[0]}/{pairPrefix}";
                return true;
            }

            if (pairError is not null)
            {
                errors.Add(pairError);
                return false;
            }
        }

        if (IpRegex().IsMatch(segment))
        {
            if (TryMaskToPrefix(segment, out _, out _))
            {
                errors.Add($"Invalid target: {segment} (subnet mask is not a scannable host address).");
                return false;
            }

            normalized = segment;
            return true;
        }

        return false;
    }

    internal static bool TryMaskToPrefix(string maskPart, out int prefix, out string? error)
    {
        prefix = 0;
        error = null;
        maskPart = maskPart.Trim();
        if (string.IsNullOrWhiteSpace(maskPart))
        {
            error = "Subnet mask or prefix length is required.";
            return false;
        }

        if (int.TryParse(maskPart, out prefix) && prefix is >= 0 and <= 32)
            return true;

        if (!DottedMaskRegex.IsMatch(maskPart))
        {
            error = maskPart.Contains('.', StringComparison.Ordinal)
                ? $"Invalid subnet mask: {maskPart} (expected four octets, e.g. 255.255.255.0)."
                : $"Invalid prefix length: {maskPart} (use 0-32 or a dotted mask like 255.255.255.0).";
            return false;
        }

        if (!IPAddress.TryParse(maskPart, out var maskIp))
        {
            error = $"Invalid subnet mask: {maskPart}.";
            return false;
        }

        prefix = SubnetMaskToPrefix(maskIp);
        if (prefix >= 0)
            return true;

        error = $"Invalid subnet mask: {maskPart} (mask bits must be contiguous).";
        return false;
    }

    internal static int SubnetMaskToPrefix(IPAddress mask)
    {
        var value = IpToUInt(mask);
        if (value == 0)
            return 0;

        var prefix = 0;
        for (var bit = 31; bit >= 0; bit--)
        {
            if ((value & (1u << bit)) == 0)
                break;
            prefix++;
        }

        var expected = prefix >= 32 ? uint.MaxValue : uint.MaxValue << (32 - prefix);
        return value == expected ? prefix : -1;
    }

    private static bool IsCidr(string token) => CidrRegex().IsMatch(token.Trim());

    private static bool IsIp(string token) => IpRegex().IsMatch(token.Trim());

    private static uint GetNetworkLong(IPAddress ip, int prefix)
    {
        var ipLong = IpToUInt(ip);
        var mask = prefix >= 32 ? uint.MaxValue : uint.MaxValue << (32 - prefix);
        return ipLong & mask;
    }

    private static uint GetBroadcastLong(uint networkLong, int prefix)
    {
        var mask = prefix >= 32 ? uint.MaxValue : uint.MaxValue << (32 - prefix);
        return networkLong | ~mask;
    }

    private static uint IpToUInt(IPAddress ip)
    {
        var bytes = ip.GetAddressBytes();
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToUInt32(bytes, 0);
    }

    private static bool TryParseIpToUInt(string ip, out uint value)
    {
        value = 0;
        return IPAddress.TryParse(ip, out var parsed) && (value = IpToUInt(parsed)) >= 0;
    }

    private static string UIntToIp(uint value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return new IPAddress(bytes).ToString();
    }

    private static uint IpSortKey(string ip) =>
        IPAddress.TryParse(ip, out var parsed) ? IpToUInt(parsed) : uint.MaxValue;

    private static string IpSortKeyToString(string ip) => ip;
}
