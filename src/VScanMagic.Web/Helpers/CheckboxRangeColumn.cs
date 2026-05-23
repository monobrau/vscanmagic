using Microsoft.AspNetCore.Components.Web;

namespace VScanMagic.Web.Helpers;

/// <summary>
/// Shift-click range selection for a vertical list of checkboxes in one column.
/// </summary>
public sealed class CheckboxRangeColumn
{
    private int? _anchorIndex;
    private bool _anchorChecked;

    public void Click(int index, bool newChecked, MouseEventArgs e, Action<int, bool> setChecked)
    {
        if (e.ShiftKey && _anchorIndex is int anchor)
        {
            var start = Math.Min(anchor, index);
            var end = Math.Max(anchor, index);
            for (var i = start; i <= end; i++)
                setChecked(i, _anchorChecked);
        }
        else
        {
            setChecked(index, newChecked);
            _anchorIndex = index;
            _anchorChecked = newChecked;
        }
    }

    public void Reset()
    {
        _anchorIndex = null;
    }
}
