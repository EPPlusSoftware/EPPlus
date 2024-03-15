/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Table.PivotTable
{
    internal static class PivotTableEnumExtentions
    {
        internal static ePivotAreaType ToPivotAreaType(this string value)
        {
            if(value == "button")
            {
                return ePivotAreaType.FieldButton;
            }
            if(value=="topRight")
            {
                return ePivotAreaType.TopEnd;
            }
            else
            {
                return value.ToEnum(ePivotAreaType.Normal);
            }
        }
        internal static string ToPivotAreaTypeString(this ePivotAreaType value)
        {
            if (value == ePivotAreaType.FieldButton)
            {
                return "button";
            }
            else
            {
                return value.ToEnumString();
            }
        }

        internal static ePivotTableAxis ToPivotTableAxis(this string value)
        {
            switch(value)
            {
                case "axisCol": 
                    return ePivotTableAxis.ColumnAxis;
                case "axisRow":
                    return ePivotTableAxis.RowAxis;
                case "axisPage":
                    return ePivotTableAxis.PageAxis;
                case "axisValues":
                    return ePivotTableAxis.ValuesAxis;
                default:
                    return ePivotTableAxis.None;
            }
        }
        internal static string ToPivotTableAxisString(this ePivotTableAxis value)
        {
            switch (value)
            {
                case ePivotTableAxis.ColumnAxis:
                    return "axisCol";
                case ePivotTableAxis.RowAxis:
                    return "axisRow";
                case ePivotTableAxis.PageAxis:
                    return "axisPage";
                case ePivotTableAxis.ValuesAxis:
                    return "axisValues";
                default:
                    return "";
            }
        }
        internal static string FromShowDataAs(this eShowDataAs value)
        {
            string text = value.ToString();
            switch (value)
            {
                case eShowDataAs.PercentDifference:
                    return "percentDiff";
                case eShowDataAs.PercentOfColumn:
                    return "percentOfCol";
                case eShowDataAs.PercentOfParentColumn:
                    return "percentOfParentCol";
                case eShowDataAs.RunningTotal:
                    return "runTotal";
                default:
                    return value.ToEnumString();
            }
        }
        internal static eShowDataAs ToShowDataAs(this string text)
        {
            switch (text)
            {
                case "percentDiff":
                    return eShowDataAs.PercentDifference;
                case "percentOfCol":
                    return eShowDataAs.PercentOfColumn;
                case "percentOfParentCol":
                    return eShowDataAs.PercentOfParentColumn;
                case "runTotal":
                    return eShowDataAs.RunningTotal;
                default:
                    return text.ToEnum(eShowDataAs.Normal);
            }
        }

    }
}