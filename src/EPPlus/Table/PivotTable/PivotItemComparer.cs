/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    public partial class ExcelPivotTableFieldItemsCollection
    {
        internal class PivotItemComparer : IComparer<ExcelPivotTableFieldItem>
		{
			private int _mult;
			private ExcelPivotTableField _field;
            private bool _hasGrouping;
			public PivotItemComparer(eSortType sort, ExcelPivotTableField field)
			{
				this._mult = sort==eSortType.Ascending ? 1 : -1;
				this._field = field;
                _hasGrouping = _field.Grouping != null;
			}

			public int Compare(ExcelPivotTableFieldItem x, ExcelPivotTableFieldItem y)
			{
                if (x.Type == eItemType.Data && y.Type == eItemType.Data)
                {
                    if(x.Value == null) return 1;
                    if(y.Value == null) return -1;
                    var xText = GetTextValue(x);
                    var yText = GetTextValue(y);
                    return xText.CompareTo(yText) * _mult;
                }
                else
                {
					return x.Type == eItemType.Data ? -1 : 1;
				}
			}

			private string GetTextValue(ExcelPivotTableFieldItem item)
			{
				if(string.IsNullOrEmpty(item.Text))
                {
					return ExcelPivotTableCacheField.GetSharedStringText(item.Value, out _);
                }
                else
                {
                    return item.Text;
                }
			}
		}
	}
}