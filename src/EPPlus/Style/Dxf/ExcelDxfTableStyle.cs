/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Xml;
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Differential formatting record used for table styles
    /// </summary>
    public class ExcelDxfTableStyle : ExcelDxfStyleLimitedFont
    {
        internal ExcelDxfTableStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) 
            : this(nameSpaceManager,topNode, styles, null)
        {
        }
        internal ExcelDxfTableStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(nameSpaceManager, topNode, styles, callback)
        {

        }
    }
}
