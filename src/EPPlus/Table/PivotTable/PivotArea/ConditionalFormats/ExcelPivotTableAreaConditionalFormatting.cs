/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/30/2024         EPPlus Software AB       Pivot Table Conditional Formatting - EPPlus 7.4
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines an area for conditional formatting within a pivot table.
    /// </summary>
    public class ExcelPivotTableAreaConditionalFormatting : ExcelPivotArea
    {
        internal ExcelPivotTableAreaConditionalFormatting(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) :
            base(nsm, topNode, pt)
        {

        }
    }
}