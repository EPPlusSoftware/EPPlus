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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.Chart;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    ///  Defines a Theme override for a chart
    /// </summary>
    public class ExcelThemeOverride : ExcelThemeBase
    {
        ExcelChartBase _chart;

        internal ExcelThemeOverride(ExcelChartBase chart, ZipPackageRelationship rel)
            : base(chart._drawings._package, chart.NameSpaceManager, rel,"")
        {
            _chart = chart;
        }
    }
}
