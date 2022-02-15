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
using System;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents an Excel Chartsheet and provides access to its properties and methods
    /// </summary>
    public class ExcelChartsheet : ExcelWorksheet
    {
        //ExcelDrawings draws;
        internal ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden? hidden, eChartType chartType, ExcelPivotTable pivotTableSource ) :
            base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
        {
            IsChartSheet = true;
            Drawings.AddAllChartTypes("Chart 1", chartType, pivotTableSource, eEditAs.Absolute);
        }
        internal ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden? hidden) :
            base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
        {
            IsChartSheet = true;
        }
        /// <summary>
        /// The worksheet chart object
        /// </summary>
        public ExcelChart Chart 
        {
            get
            {
                return (ExcelChart)Drawings[0];
            }
        }
    }
}
