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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;

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
            BoundingBox worldBounds = new BoundingBox();
            worldBounds.Width = 400;
            worldBounds.Height = 400;

            worldBounds.transform.Name = "World Bounds";

            BoundingBox shapeRect = new BoundingBox();

            shapeRect.Parent = worldBounds;

            shapeRect.Width = 200;
            shapeRect.Height = 200;

            shapeRect.X = 20;
            shapeRect.Y = 20;

            shapeRect.transform.Name = "Shape";

            FontMeasurerTrueType measurer = new FontMeasurerTrueType(11, "Aptos Narrow", FontSubFamily.Regular);
            var body = new TextBody(shapeRect);

            body.Bounds.transform.Name = "TxtBody";

            body.Bounds.X = 20;
            body.Bounds.Y = 20;

            body.Bounds.Width = 100;
            body.Bounds.Height = 100;

            body.AddText(txtBody, measurer);

            var para1 = body.Paragraphs[0];

            para1.AddText("Extra Text", measurer);
            para1.Runs[1].X = 10;
            para1.Runs[1].Y = 20;

            para1.Bounds.Width = 120;
            para1.Bounds.Height = 100;

            //body.AddParagraph("Paragraph2 text", measurer);
            //var para2 = body.Paragraphs[1];

            //para2.Bounds.Y = 40;

            //para2.AddText("Para2 Run2", measurer);
            //para2.Runs[1].X = 5;
            //para2.Runs[1].Y = 20;

            var svgBody = GenerateSvgTextBody(body, (int)worldBounds.Width, (int)worldBounds.Height);

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

            //Represents world bounds/svg node
            var bg = new SvgElement("rect");
            bg.AddAttribute("width", "100%");
            bg.AddAttribute("height", "100%");
            bg.AddAttribute("fill", "red");
            bg.AddAttribute("opacity", "0.1");

            body.AllowOverflow = false;

            var svgDefs = GetDefinitions(body.Bounds, out string nameId, body.AllowOverflow);

            var fontSizePx = 16d;

            doc.AddChildElement(svgDefs);
            doc.AddChildElement(bg);

            var shapeRectBB = body.Bounds.Parent;

            var shapeRoot = new SvgElement("g");
            shapeRoot.AddAttribute("transform", $"translate({shapeRectBB.GlobalX},{shapeRectBB.GlobalY})");

            doc.AddChildElement(shapeRoot);

            var shapeTitle = new SvgElement("title");
            shapeTitle.Content = "Shape Group";
            shapeRoot.AddChildElement(shapeTitle);

            var shapeVisual = new SvgElement("rect");
            shapeVisual.AddAttribute("width", $"{shapeRectBB.Width}px");
            shapeVisual.AddAttribute("height", $"{shapeRectBB.Height}px");
            shapeVisual.AddAttribute("fill", "yellow");
            shapeVisual.AddAttribute("opacity", "0.2");

            shapeRoot.AddChildElement(shapeVisual);

            var textBodyGroup = new SvgElement("g");
            textBodyGroup.AddAttribute("transform", $"translate({body.Bounds.X},{body.Bounds.Y})");

            shapeRoot.AddChildElement(textBodyGroup);

            var txtBodyTitle = new SvgElement("title");
            txtBodyTitle.Content = "txtBody";
            textBodyGroup.AddChildElement(txtBodyTitle);


            var txBodyVisual = new SvgElement("use");
            txBodyVisual.AddAttribute("href", "#defaultRect");
            txBodyVisual.AddAttribute("fill", "green");
            txBodyVisual.AddAttribute("opacity", "0.5");

            textBodyGroup.AddChildElement(txBodyVisual);

            int paragraphCount = 1;

            foreach(var paragraph in body.Paragraphs)
            {
                var paragraphGroup = new SvgElement("g");
                paragraphGroup.AddAttribute("transform", $"translate({paragraph.Bounds.X},{paragraph.Bounds.Y})");

                textBodyGroup.AddChildElement(paragraphGroup);

                var paragraphTitle = new SvgElement("title");
                paragraphTitle.Content = "Paragraph " + paragraphCount.ToString();
                paragraphGroup.AddChildElement(paragraphTitle);

                var paragraphElement = new SvgElement("text");
                paragraphElement.AddAttribute("y", fontSizePx);
                paragraphElement.AddAttribute("_measurementFont-size", $"{fontSizePx}px");
                paragraphElement.AddAttribute("clip-path", $"url(#{nameId})");

                paragraphGroup.AddChildElement(paragraphElement);

                foreach (var run in paragraph.Runs)
                {
                    var bbVisual = new SvgElement("rect");
                    bbVisual.AddAttribute("x", run.X);
                    bbVisual.AddAttribute("y", run.Y);
                    bbVisual.AddAttribute("width", run.Width);
                    bbVisual.AddAttribute("height", run.Height);
                    bbVisual.AddAttribute("fill", "blue");
                    bbVisual.AddAttribute("opacity", "0.5");

                    paragraphGroup.AddChildElement(bbVisual);

                    var runElement = new SvgElement("tspan");
                    runElement.AddAttribute("x", run.X);
                    runElement.AddAttribute("y", run.Y + fontSizePx);
                    runElement.AddAttribute("_measurementFont-size", $"{fontSizePx}px");

                    runElement.Content = run.GetContent();
                    paragraphElement.AddChildElement(runElement);
                }

                paragraphCount++;
            }

            doc.AddAttributes();

            return doc;
        }

        internal SvgElement GetDefinitions(BoundingBox boundingBox, out string nameId, bool AllowOverflow = false)
        {
            nameId = "boundingBox";
            var def = new SvgElement("defs");

            string defaultName = "defaultRect";

            if (AllowOverflow == false)
            {
                var bb = new SvgElement("rect");
                bb.AddAttribute("width", boundingBox.Width);
                bb.AddAttribute("height", boundingBox.Height);
                bb.AddAttribute("id", defaultName);

                def.AddChildElement(bb);

                var clipPath = new SvgElement("clipPath");
                clipPath.AddAttribute("id", nameId);

                def.AddChildElement(clipPath);

                var useElement = new SvgElement("use");
                useElement.AddAttribute("href", $"#{defaultName}");

                clipPath.AddChildElement(useElement);
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
            renderElement.AddAttribute("_measurementFont-size", $"{fontSizePx}px");
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
