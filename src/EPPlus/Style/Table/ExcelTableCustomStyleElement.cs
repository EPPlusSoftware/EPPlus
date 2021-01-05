using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;

namespace OfficeOpenXml.Style
{
    public class ExcelTableCustomStyleElement 
    {
        ExcelStyles _styles;
        XmlNamespaceManager _nsm;
        XmlNode _topNode;
        internal ExcelTableCustomStyleElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles)
        {
            _styles = styles;
            _nsm = nameSpaceManager;
            _topNode = topNode;
        }
        public int DxfId
        {
            get;
            internal set;
        }
        ExcelDxfStyleConditionalFormatting _style = null;
        public ExcelDxfStyleConditionalFormatting Style
        {
            get
            {
                if (_style == null)
                {
                    _style = new ExcelDxfStyleConditionalFormatting(_nsm, _topNode, _styles);
                }
                return _style;
            }
        }
        int _bandSize = 1;
        /// <summary>
        /// Band size. Only applicable when <see cref="Type"/> is set to FirstRowStripe, FirstColumnStripe, SecondRowStripe or SecondColumnStripe
        /// </summary>
        public int BandSize
        {
            get
            {
                return _bandSize;
            }
            set
            {
                if(value < 1 && value > 9)
                {
                    throw new InvalidOperationException("BandSize must be between 1 and 9");
                }
                _bandSize = value;
            }
        }
        public eTableStyleElement Type
        {
            get;
            set;
        }
    }
}
