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
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for table style element
    /// </summary>
    public class ExcelTableStyleElementXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelTableStyleElementXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            _styles = styles;
        }
        internal ExcelTableStyleElementXml(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(nameSpaceManager, topNode)
        {
            Type = GetXmlNodeString(typePath);
            DxfId = GetXmlNodeInt(dxfIdPath);
            Size = GetXmlNodeInt(sizePath);

            _styles = styles;
        }
        internal override string Id
        {
            get
            {
                return Type;
            }
        }
        const string dxfIdPath = "@dxfId";
        /// <summary>
        /// Zero-based index to a dxf record in the dxfs collection, specifying differential formatting to use with this Table or PivotTable style element.
        /// </summary>
        public int DxfId { get; set; }

        const string sizePath = "@size";
        /// <summary>
        /// Count of table style elements defined for this table style
        /// </summary>
        public int Size { get; set; }

        const string typePath = "@type";

        /// <summary>
        /// Name of the style
        /// </summary>
        public string Type { get; internal set; }

        public ExcelDxfStyleConditionalFormatting DxfStyle
        {
            get { return _styles.Dxfs[DxfId]; }
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString(typePath, Type);
            SetXmlNodeInt(dxfIdPath, DxfId);
            if (Size > 0) SetXmlNodeInt(sizePath, Size);

            return TopNode;            
        }
    }
}
