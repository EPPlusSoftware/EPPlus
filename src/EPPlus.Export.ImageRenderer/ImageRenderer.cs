/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlus.Export.ImageRenderer;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Export.ImageRenderer.Svg.Writer;
using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace EPPlusImageRenderer
{
    public class ImageRenderer
    {
        public string RenderDrawingToSvg(ExcelDrawing drawing)
        {
            drawing.GetSizeInPixels(out int width, out int height);
            var sb = new StringBuilder();
            if (drawing is ExcelShape shape)
            {
                var svg = new SvgShape(shape);
                svg.Size = new DrawingSize(width, height);
                svg.Render(sb);
                return sb.ToString();
            }
            else if(drawing is ExcelChart chart)
            {
                var svg = new SvgChart(chart);
                svg.Size = new DrawingSize(width, height);
                svg.Render(sb);
                return sb.ToString();
            }

            throw new NotImplementedException("Image rendering for drawing type not implemented.");
        }
        public string RenderRangeToSvg(ExcelRange range)
        {
            var ws = range.Worksheet;

            //ws.Workbook.Styles.CellXfs.

            double totalWidth = 0;
            foreach (var col in range.EntireColumn)
            {
                totalWidth += ExcelColumn.ColumnWidthToPixels(col.Width, range.Worksheet.Workbook.MaxFontWidth);
            }
            double totalHeight = 0;
            foreach (var row in range.EntireRow)
            {
                totalHeight = row.Height;
            }

            var sRange = new SvgRange(range, totalWidth, totalHeight);

            var sb = new StringBuilder();
            sb.Append($"<svg width=\"{500}\" height=\"{500}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\">");
            sRange.Render(sb);
            sb.Append($"</svg>");
            return sb.ToString();
        }

        public string RenderBox(string boxText)
        {
            string retStr = "";
            var container = new TextContainerBase(boxText);
            var element = GenerateSvg(container);

            using (var ms = EPPlusMemoryManager.GetStream())
            {
                SvgWriter writer = new SvgWriter(ms, Encoding.UTF8);
                writer.RenderSvgElement(element, true);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    retStr = sr.ReadToEnd();
                    return retStr;
                }
            }

            //writer.RenderSvgElement(element, true);

            //StreamReader reader = new StreamReader(ms);
            //retStr = reader.ReadToEnd();
            
            ////SvgParagraph para = new SvgParagraph(container.GetContent(),);
            ////var doc = new SvgEpplusDocument();
            //return retStr;
        }

        public string RenderTextBody(string txtBody)
        {
            BoundingBox shapeRect = new BoundingBox();

            shapeRect.Width = 200;
            shapeRect.Height = 200;

            FontMeasurerTrueType measurer = new FontMeasurerTrueType(11, "Aptos Narrow", FontSubFamily.Regular);
            var body = new TextBody(shapeRect);

            body.Bounds.transform.Name = "TxtBody";

            body.Bounds.X = 40;
            body.Bounds.Y = 40;

            body.AddText(txtBody, measurer);

            var para1 = body.Paragraphs[0];

            para1.Runs[0].X = 20;
            para1.Runs[0].Y = 20;

            para1.AddText("Extra Text", measurer);
            para1.Runs[1].X = 30;
            para1.Runs[1].Y = 40;

            para1.Bounds.Width = 120;
            para1.Bounds.Height = 100;

            var svgBody = GenerateSvgTextBody(body, (int)shapeRect.Width, (int)shapeRect.Height);

            return RenderSvgElement(svgBody);
        }

        internal string RenderSvgElement(SvgElement element)
        {
            string retStr = string.Empty;

            using (var ms = EPPlusMemoryManager.GetStream())
            {
                SvgWriter writer = new SvgWriter(ms, Encoding.UTF8);
                writer.RenderSvgElement(element, true);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    retStr = sr.ReadToEnd();
                    return retStr;
                }
            }
        }

        internal SvgElement GenerateSvgTextBody(TextBody body, int width, int height)
        {
            var fullString = body.GetContent();

            var doc = new SvgEpplusDocument(width, height);

            var bg = new SvgElement("rect");
            bg.AddAttribute("width", "100%");
            bg.AddAttribute("height", "100%");
            bg.AddAttribute("fill", "red");
            bg.AddAttribute("opacity", "0.1");

            body.AllowOverflow = true;

            var svgDefs = GetDefinitions(body.Bounds, out string nameId, body.AllowOverflow);

            var fontSizePx = 16d;

            doc.AddChildElement(svgDefs);
            doc.AddChildElement(bg);

            foreach (var run in body.Paragraphs[0].Runs)
            {
                var bbVisual = new SvgElement("rect");
                bbVisual.AddAttribute("x", run.GlobalX);
                bbVisual.AddAttribute("y", run.GlobalY);
                bbVisual.AddAttribute("width", run.Width);
                bbVisual.AddAttribute("height", run.Height);
                bbVisual.AddAttribute("fill", "blue");
                bbVisual.AddAttribute("opacity", "0.5");

                doc.AddChildElement(bbVisual);
            }

            var txBodyVisual = new SvgElement("rect");
            txBodyVisual.AddAttribute("x", body.Bounds.GlobalX);
            txBodyVisual.AddAttribute("y", body.Bounds.GlobalY);
            txBodyVisual.AddAttribute("width", body.Bounds.Width);
            txBodyVisual.AddAttribute("height", body.Bounds.Height);
            txBodyVisual.AddAttribute("fill", "green");
            txBodyVisual.AddAttribute("opacity", "0.5");

            doc.AddChildElement(txBodyVisual);

            foreach (var run in body.Paragraphs)
            {
                var renderElement = new SvgElement("text");
                renderElement.AddAttribute("x", run.Bounds.GlobalX);
                renderElement.AddAttribute("y", run.Bounds.GlobalY + fontSizePx);
                renderElement.AddAttribute("font-size", $"{fontSizePx}px");
                renderElement.AddAttribute("clip-path", $"url(#{nameId})");

                renderElement.Content = run.GetContent();
                doc.AddChildElement(renderElement);
            }

            doc.AddAttributes();

            return doc;
        }

        internal SvgElement GetDefinitions(BoundingBox boundingBox, out string nameId, bool AllowOverflow = false)
        {
            nameId = "boundingBox";
            var def = new SvgElement("defs");

            if (AllowOverflow == false)
            {
                var clipPath = new SvgElement("clipPath");
                clipPath.AddAttribute("id", nameId);

                def.AddChildElement(clipPath);

                var bb = new SvgElement("rect");
                bb.AddAttribute("x", boundingBox.GlobalX);
                bb.AddAttribute("y", boundingBox.GlobalY);
                bb.AddAttribute("width", boundingBox.Width);
                bb.AddAttribute("height", boundingBox.Height);

                clipPath.AddChildElement(bb);
            }

            return def;
        }

        internal SvgElement GenerateSvg(TextContainerBase container)
        {
            var fullString = container.GetContent();

            var doc = new SvgEpplusDocument(500, 500);

            var bg = new SvgElement("rect");
            bg.AddAttribute("width", "100%");
            bg.AddAttribute("height", "100%");
            bg.AddAttribute("fill", "red");
            bg.AddAttribute("opacity", "0.1");

            var nameId = "boundingBox";
            var def = new SvgElement("defs");
            var clipPath = new SvgElement("clipPath");
            clipPath.AddAttribute("id", nameId);

            def.AddChildElement(clipPath);

            var bb = new SvgElement("rect");
            bb.AddAttribute("x", container.transform.Position.X);
            bb.AddAttribute("y", container.transform.Position.Y);
            bb.AddAttribute("width", container.Width);
            bb.AddAttribute("height", container.Height);
            //bb.AddAttribute("fill", "blue");
            //bb.AddAttribute("opacity", "0.5");

            clipPath.AddChildElement(bb);

            var fontSizePx = 16d;

            var renderElement = new SvgElement("text");
            renderElement.AddAttribute("x", container.transform.Position.X);
            renderElement.AddAttribute("y", container.transform.Position.Y + fontSizePx);
            renderElement.AddAttribute("font-size", $"{fontSizePx}px");
            renderElement.AddAttribute("clip-path", $"url(#{nameId})");

            renderElement.Content = fullString;

            var bbVisual = new SvgElement("rect");
            bbVisual.AddAttribute("x", container.transform.Position.X);
            bbVisual.AddAttribute("y", container.transform.Position.Y);
            bbVisual.AddAttribute("width", container.Width);
            bbVisual.AddAttribute("height", container.Height);
            bbVisual.AddAttribute("fill", "blue");
            bbVisual.AddAttribute("opacity", "0.5");

            doc.AddChildElement(def);
            doc.AddChildElement(bg);
            doc.AddChildElement(bbVisual);
            doc.AddChildElement(renderElement);

            doc.AddAttributes();

            return doc;
        }
    }
}
