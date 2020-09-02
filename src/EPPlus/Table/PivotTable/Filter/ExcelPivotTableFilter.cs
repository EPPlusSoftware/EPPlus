/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extentions;
using System.Xml;

namespace EPPlusTest.Table.PivotTable.Filter
{
    public class ExcelPivotTableFilter : XmlHelper
    {
        internal ExcelPivotTableFilter(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        public int Id
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            internal set
            {
                SetXmlNodeInt("@id", value);
            }
        }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value, true);
            }
        }
        public string Description
        {
            get
            {
                return GetXmlNodeString("@description");
            }
            set
            {
                SetXmlNodeString("@description", value, true);
            }
        }
        public ePivotTableFilterType Type
        {
            get
            {
                return GetXmlNodeString("@type").ToEnum(ePivotTableFilterType.Unknown);
            }
            internal set
            {
                SetXmlNodeString("@type", value.ToEnumString());
            }
        }
        public int EvalOrder
        {
            get
            {
                return GetXmlNodeInt("@evalOrder");
            }
            internal set
            {
                SetXmlNodeInt("@evalOrder", value);
            }
        }
    }
}
