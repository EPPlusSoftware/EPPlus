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
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for table styles
    /// </summary>
    public class ExcelTableStyleXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelTableStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            _styles = styles;
        }
        internal ExcelTableStyleXml(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(nameSpaceManager, topNode)
        {
            Name = GetXmlNodeString(namePath);
            Pivot = GetXmlNodeBool(pivotPath);
            Count = GetXmlNodeInt(countPath);
            Table = GetXmlNodeBool(tablePath);

            Elements = new ExcelStyleCollection<ExcelTableStyleElementXml>();

            foreach (XmlNode n in topNode)
            {
                ExcelTableStyleElementXml item = new ExcelTableStyleElementXml(nameSpaceManager, n, styles);
                Elements.Add(item.Id, item);
            }

            _styles = styles;
        }
        internal override string Id
        {
            get
            {
                return Name;
            }
        }
        const string pivotPath = "@pivot";
        /// <summary>
        /// 'True' if this table style should be shown as an available pivot table style.
        /// </summary>
        public bool Pivot { get; set; }

        const string countPath = "@count";
        /// <summary>
        /// Count of table style elements defined for this table style
        /// </summary>
        public int Count { get; set; }

        const string tablePath = "@table";
        /// <summary>
        /// True if this table style should be shown as an available table style.
        /// </summary>
        public bool Table { get; set; }

        const string namePath = "@name";

        /// <summary>
        /// Name of the style
        /// </summary>
        public string Name { get; internal set; }

        private const string elementPath = "tableStyleElement";
        public ExcelStyleCollection<ExcelTableStyleElementXml> Elements { get; set; }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString(namePath, Name);
            SetXmlNodeBool(pivotPath, Pivot);
            SetXmlNodeInt(countPath, Count);
            if (Table) SetXmlNodeBool(tablePath, Table);
            return TopNode;            
        }
    }
}
