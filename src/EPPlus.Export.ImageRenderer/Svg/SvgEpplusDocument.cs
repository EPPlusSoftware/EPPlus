using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Export.ImageRenderer.Svg.Nodes;
using EPPlus.Export.ImageRenderer.Svg.Writer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg
{
    internal class SvgEpplusDocument : SvgElement
    {
        const string svgNamespaceValue = "http://www.w3.org/2000/svg";
        const string xLinkNameSpaceValue = "http://www.w3.org/1999/xlink";

        //const string xmlNameSpace = "xmlns=\"http://www.w3.org/2000/svg\"";
        //const string xLink = "xmlns:xlink=\"http://www.w3.org/1999/xlink\"";
        bool preserveWhiteSpace = false;
        bool useViewBox = true;
        OverflowAttribute Overflow = null;

        internal DrawingSize SvgSize;

        public string ViewBox
        {
            get
            {
                double l = 0, t = 0, r = 1, b = 1;
                //TODO: Insert Calculate bounds from render items here

                return $"{(l * SvgSize.Width).ToString(CultureInfo.InvariantCulture)},{(t * SvgSize.Height).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(l) + r) * SvgSize.Width).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(t) + b) * SvgSize.Height).ToString(CultureInfo.InvariantCulture)}";
            }
        }

        internal SvgEpplusDocument(int width, int height) : base("svg")
        {
            SvgSize = new DrawingSize(width, height);
            Overflow = new OverflowAttribute();
            Overflow.OverFlowValue = eOverFlowValues.Hidden;
        }

        internal void AddAttributes()
        {
            AddAttribute("width", SvgSize.Width);
            AddAttribute("height", SvgSize.Height);

            AddAttribute("xmlns", svgNamespaceValue);
            AddAttribute("xmlns:xlink", xLinkNameSpaceValue);

            AddAttribute("xml:space", preserveWhiteSpace ? "preserve" : "default");

            if (Overflow != null)
            {
                AddAttribute(Overflow.Name, Overflow.Value);
            }

            if (useViewBox)
            {
                AddAttribute("viewBox", ViewBox);
            }
        }

        internal void Render(MemoryStream stream)
        {
            AddAttributes();
            SvgWriter writer = new SvgWriter(stream, Encoding.UTF8);
            writer.RenderSvgElement(this, true);
        }
    }
}
