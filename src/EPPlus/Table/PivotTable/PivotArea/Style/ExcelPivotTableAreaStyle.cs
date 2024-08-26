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
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table area of selection used for styling.
    /// </summary>
    public class ExcelPivotTableAreaStyle : ExcelPivotArea
    {
        ExcelStyles _styles;
        internal ExcelPivotTableAreaStyle(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) :
            base(nsm, topNode, pt)
        {
            _styles = pt.WorkSheet.Workbook.Styles;
            Conditions = new ExcelPivotAreaStyleConditions(nsm, topNode, pt);
        }
        /// <summary>
        /// Conditions for the pivot table. Conditions can be set for specific row-, column- or data fields. Specify labels, data grand totals and more.
        /// </summary>
        public ExcelPivotAreaStyleConditions Conditions
        {
            get;
        }

        ExcelDxfStyle _style = null;
        /// <summary>
        /// Access to the style property for the pivot area
        /// </summary>
        public ExcelDxfStyle Style 
        { 
            get
            {
                if (_style == null)
                {
                    var dxfId= GetXmlNodeIntNull("../@dxfId");
                    _style = _styles.GetDxf(dxfId, null);
                }
                return _style;
            }
            internal set
            {
                _style = value;
            }
        }

        internal int? DxfId 
        { 
            get
            {
                return GetXmlNodeIntNull("../@dxfId");
            }
            set
            {
                SetXmlNodeInt("../@dxfId", value);
            }
        }

    }
}
