/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.7
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines sorting for a pivot table area within a pivot table.
    /// </summary>
    public class ExcelPivotAreaAutoSort : ExcelPivotArea
    {
        internal ExcelPivotAreaAutoSort(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) :
            base(nsm, topNode, pt)
        {

        }
    }
}