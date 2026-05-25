namespace VScanMagic.ConnectSecure;

internal static class ConnectSecureSuppressQueryExtensions
{
    public static Dictionary<string, string> WithLookupLimit(this Dictionary<string, string> query, int limit = 50)
    {
        query["limit"] = limit.ToString();
        query["skip"] = "0";
        return query;
    }

    public static Dictionary<string, string> WithFilter(this Dictionary<string, string> query, string filter)
    {
        query["filter"] = filter;
        return query;
    }
}
