/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.3
 ******0*******************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicer : ExcelSlicer<ExcelPivotTableSlicerCache>
    {
        internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {

        }
        public ExcelPivotTableCollection PivotTables
        {
            get;
        } = new ExcelPivotTableCollection();
    }
}
