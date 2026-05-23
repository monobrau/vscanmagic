using System.Net;
using System.Text.RegularExpressions;

namespace VScanMagic.ConnectSecure;

public static partial class ExternalSubnetHelper
{
    [GeneratedRegex(@"^(\d+\.\d+\.\d+\.\d+)/(\d+)$")]
    private static partial Regex CidrRegex();

    [GeneratedRegex(@"^\d+\.\d+\.\d+\.\d+$")]
    private static partial Regex IpRegex();

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
        var tokens = TokenizeInput(input);
        if (tokens.Count == 0)
        {
            return new ExternalScanTargetValidationResult(
                false,
                ["Enter a CIDR subnet or one or more IP addresses."],
                [],
                "",
                "");
        }

        var cidrTokens = tokens.Where(IsCidr).ToList();
        var ipTokens = tokens.Where(t => IsIp(t) && !IsCidr(t)).ToList();
        var invalidTokens = tokens.Where(t => !IsCidr(t) && !IsIp(t)).ToList();

        var errors = new List<string>();
        foreach (var invalid in invalidTokens)
            errors.Add($"Invalid target: {invalid}");

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

    private static List<string> TokenizeInput(string input) =>
        input
            .Split([',', ';', '\n', '\r', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();

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
