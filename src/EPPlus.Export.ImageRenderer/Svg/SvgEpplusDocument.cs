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

        const string xmlNameSpace = "xmlns=\"http://www.w3.org/2000/svg\"";
        const string xLink = "xmlns:xlink=\"http://www.w3.org/1999/xlink\"";
        bool preserveWhiteSpace = true;
        bool useViewBox = true;
        OverflowAttribute Overflow = null;

        internal DrawingSize Size;

        public string ViewBox
        {
            get
            {
                double l = 0, t = 0, r = 1, b = 1;
                //TODO: Insert Calculate bounds from render items here

                return $"{(l * Size.Width).ToString(CultureInfo.InvariantCulture)},{(t * Size.Height).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(l) + r) * Size.Width).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(t) + b) * Size.Height).ToString(CultureInfo.InvariantCulture)}";
            }
        }

        internal SvgEpplusDocument() : base("svg")
        {
            Overflow = new OverflowAttribute();
        }

        internal void AddAttributes()
        {
            AddAttribute("width", Size.Width);
            AddAttribute("height", Size.Height);

            AddAttribute("xmlns", svgNamespaceValue);
            AddAttribute("xmlns:xlink", xLinkNameSpaceValue);

            AddAttribute("xml:space=", preserveWhiteSpace ? "preserve" : "default");

            if (Overflow != null)
            {
                AddAttribute(Overflow.Name, Overflow.Value);
            }

            if (useViewBox)
            {
                AddAttribute("viewBox", ViewBox);
            }
        }

        internal void RenderStartNode(StringBuilder sb)
        {
            string attributes = $"width=\"{Size.Width}\" height=\"{Size.Height}\" {xmlNameSpace} {xLink} {GetXmlSpace()}";

            if (Overflow != null)
            {
                attributes += Overflow.Render();
            }

            if(useViewBox)
            {
                attributes += $" viewbox=\"{ViewBox}\" ";
            }

            sb.Append($"<svg {attributes}>");
        }

        internal void RenderEndNode(StringBuilder sb)
        {
            sb.AppendLine("</svg>");
        }

        internal List<SvgElement> _childElements = new List<SvgElement>();

        internal void Render(MemoryStream stream)
        {
            AddAttributes();
            SvgWriter writer = new SvgWriter(stream, Encoding.UTF8);
            writer.RenderSvgElement(this, true);

            //attr


            //StringBuilder sb = new StringBuilder();
            //var svgWriter = new SvgWriter(stream, Encoding.UTF8);
            //RenderStartNode(sb);
            //foreach (SvgElement element in _childElements)
            //{
            //    element.
            //}
        }

        string GetXmlSpace()
        {
            var spaceStr = $"xml:space=";
            if (preserveWhiteSpace)
            {
                spaceStr += "\"preserve\"";
            }
            else
            {
                spaceStr += "\"default\"";
            }
            return spaceStr;
        }
    }
}
