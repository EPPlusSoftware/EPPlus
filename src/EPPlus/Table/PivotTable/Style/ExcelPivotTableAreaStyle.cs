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
using OfficeOpenXml.Style.Dxf;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table area of selection used for styling.
    /// </summary>
    public class ExcelPivotTableAreaStyle : ExcelPivotArea
    {
        ExcelStyles _styles;
        internal ExcelPivotTableAreaStyle(XmlNamespaceManager nsm, XmlNode topNode, ExcelStyles styles) :
            base(nsm, topNode)
        {
            _styles = styles;
        }
        public ExcelPivotAreaReferenceCollection References
        {
            get;
        }

        ExcelDxfStyle _style = null;
        public ExcelDxfStyle Style 
        { 
            get
            {
                if (_style == null)
                {
                    _style=new ExcelDxfStyle(NameSpaceManager, TopNode, _styles, "../@dxfId");
                }
                return _style;
            }
        }
    }
}
