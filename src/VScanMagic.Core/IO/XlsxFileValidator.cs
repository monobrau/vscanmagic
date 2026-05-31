namespace VScanMagic.Core.IO;

public static class XlsxFileValidator
{
    private const int ZipEndOfCentralDirectorySignature = 0x06054b50;

    public static bool IsLikelyValidXlsx(string path)
    {
        if (!File.Exists(path))
            return false;

        var info = new FileInfo(path);
        if (info.Length < 512)
            return false;

        using var stream = File.OpenRead(path);

        Span<byte> header = stackalloc byte[2];
        if (stream.Read(header) != 2 || header[0] != 0x50 || header[1] != 0x4B)
            return false;

        return HasZipEndOfCentralDirectory(stream, info.Length);
    }

    public static bool IsCorruptExcelError(Exception ex)
    {
        for (var current = ex; current is not null; current = current.InnerException)
        {
            if (current.Message.Contains("corrupted data", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("invalid signature", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("central directory", StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    public static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrEmpty(path))
            return;

        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // Best effort before retry.
        }
    }

    private static bool HasZipEndOfCentralDirectory(FileStream stream, long length)
    {
        const int maxTail = 65557;
        var tailSize = (int)Math.Min(maxTail, length);
        var tail = new byte[tailSize];
        stream.Seek(-tailSize, SeekOrigin.End);
        if (stream.Read(tail, 0, tailSize) < 22)
            return false;

        for (var i = tailSize - 22; i >= 0; i--)
        {
            if (BitConverter.ToInt32(tail, i) == ZipEndOfCentralDirectorySignature)
                return true;
        }

        return false;
    }
}
