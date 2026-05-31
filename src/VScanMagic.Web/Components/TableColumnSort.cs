namespace VScanMagic.Web.Components;

public sealed class TableColumnSort
{
    public string Column { get; private set; } = "";
    public bool Descending { get; private set; }

    public TableColumnSort(string defaultColumn, bool defaultDescending = true)
    {
        Column = defaultColumn;
        Descending = defaultDescending;
    }

    public void Toggle(string column, bool defaultDescending = true)
    {
        if (Column == column)
            Descending = !Descending;
        else
        {
            Column = column;
            Descending = defaultDescending;
        }
    }

    public string Icon(string column) =>
        Column == column ? (Descending ? "↓" : "↑") : "";
}
