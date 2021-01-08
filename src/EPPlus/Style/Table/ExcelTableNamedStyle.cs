/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Style.Table
{
    public class ExcelTableNamedStyle : ExcelTableNamedStyleBase
    {
        internal ExcelTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(nameSpaceManager, topNode, styles)
        {
        }
        public ExcelTableStyleElement LastHeaderCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastHeaderCell, false);
            }
        }
        public ExcelTableStyleElement FirstTotalCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstTotalCell, false);
            }
        }
        public ExcelTableStyleElement LastTotalCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastTotalCell, false);
            }
        }

    }
}
