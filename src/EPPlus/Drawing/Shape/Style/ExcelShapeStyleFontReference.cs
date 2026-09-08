using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Shape.Style
{
    //Represents CT_FontReference
    internal class ExcelShapeStyleFontReference : XmlHelper
    {
        string _path;
        internal ExcelShapeStyleFontReference(XmlNamespaceManager nsm, XmlNode topNode, string path) : base(nsm, topNode)
        {
            _path = path;
        }
        /// <summary>
        /// The index to the style matrix.
        /// This property referes to the theme
        /// </summary>
        public eThemeFontCollectionType Index
        {
            get
            {
                return GetXmlNodeString($"{_path}/@idx").ToEnum(eThemeFontCollectionType.None);
            }
            set
            {
                SetXmlNodeString($"{_path}/@idx", value.ToEnumString());
            }
        }
        ExcelDrawingColorManager _color = null;
        /// <summary>
        /// The color of the font
        /// This will replace any the StyleClr node in the chart style xml.
        /// </summary>
        public ExcelDrawingColorManager Color
        {
            get
            {
                if (_color == null)
                {
                    _color = new ExcelDrawingColorManager(NameSpaceManager, TopNode, _path, SchemaNodeOrder);
                }

                return _color;
            }
        }
        /// <summary>
        /// If the reference has a color
        /// </summary>
        public bool HasColor
        {
            get
            {
                var node = GetNode(_path);
                return node != null && node.HasChildNodes;
            }
        }
    }
}
