using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

//StyleMatrixReference
namespace OfficeOpenXml.Drawing.Shape.Style
{
    public class ShapeStyleReference : XmlHelper
    {
        string _path;
        internal ShapeStyleReference(XmlNamespaceManager nsm, XmlNode topNode, string path) : base(nsm, topNode)
        {
            _path = path;
        }

        /// <summary>
        /// The index to the theme style matrix.
        /// <seealso cref="ExcelWorkbook.ThemeManager"/>
        /// </summary>
        public int Index
        {
            get
            {
                return GetXmlNodeInt($"{_path}/@idx");
            }
            set
            {
                if (value < 0) throw new ArgumentOutOfRangeException("Index", "Can't be negative");
                SetXmlNodeString($"{_path}/@idx", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        ExcelDrawingColorManager _color;
        /// <summary>
        /// The color to be used for the reference. 
        /// simplerForm of Color on ChartNodes
        /// </summary>
        public ExcelDrawingColorManager ShapeColor
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
