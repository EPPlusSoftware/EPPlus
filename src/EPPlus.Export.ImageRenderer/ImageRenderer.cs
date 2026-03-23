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

using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Export.ImageRenderer.Svg.Writer;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
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
            var sb = new StringBuilder();
            if (drawing is ExcelShape shape)
            {
                var svg = new SvgShape(shape);
                svg.Render(sb);
                return sb.ToString();
            }
            else if(drawing is ExcelChart chart)
            {
                var svg = new SvgChart(chart);
                svg.Render(sb);
                return sb.ToString();
            }

            throw new NotImplementedException("Image rendering for drawing type not implemented.");
        }

        public enum RenderPresets
        {
            ContainerMargins,
            RotatingContainer
        }

        public string RenderTest(RenderPresets preset)
        {
            return GenerateFromPreset(preset);
        }


        private string rotatingContainer()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            BoundingBox parent = new BoundingBox();

            var groupItem = new SvgGroupItemNew(baseItem, parent, 45);

            groupItem.Position.Left = 10;
            groupItem.Position.Top = 10;

            SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

            rectItem.FillColor = "red";
            rectItem.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            groupItem.AddChildItem(rectItem);


            SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
            siblingItem.FillColor = "blue";
            siblingItem.FillOpacity = 0.2d;

            siblingItem.Width = 20;
            siblingItem.Height = 20;

            siblingItem.Bounds.Left = 20;
            siblingItem.Bounds.Top = 20;

            groupItem.AddChildItem(siblingItem);

            groupItem.SetRotationPointToCenterOfGroup();

            SvgRenderRectItem centerOfGroupMarker = new SvgRenderRectItem(baseItem, baseItem.Bounds);
            centerOfGroupMarker.FillColor = "green";
            centerOfGroupMarker.FillOpacity = 0.8d;

            centerOfGroupMarker.Width = 6;
            centerOfGroupMarker.Height = 6;

            centerOfGroupMarker.Left = 30 - (centerOfGroupMarker.Width / 2);
            centerOfGroupMarker.Top = 30 - (centerOfGroupMarker.Height / 2);

            baseItem.RenderItems.Add(centerOfGroupMarker);

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);

            return sb.ToString();
        }

        private string containerMargins()
        {
            var baseBB = new BoundingBox();

            baseBB.Width = 400;
            baseBB.Height = 400;

            var baseItem = new DrawingItemForTesting(baseBB);

            SvgRenderRectItem myBgItem = new SvgRenderRectItem(baseItem, baseItem.Bounds);
            myBgItem.FillColor = "purple";
            myBgItem.FillOpacity = 0.2d;

            SvgRenderRectItem myInnerItem = new SvgRenderRectItem(baseItem, myBgItem.Bounds);

            myInnerItem.FillColor = "green";
            myInnerItem.FillOpacity = 0.8d;

            myInnerItem.Width = 50;
            myInnerItem.Height = 50;

            var container = new SvgContainerItem(myInnerItem, myBgItem);

            container.MarginLeft = 5;
            container.MarginRight = 5;
            container.MarginTop = 5;
            container.MarginBottom = 5;

            container.ApplyMargins();

            baseItem.RenderItems.Add(container);

            var sb = new StringBuilder();

            baseItem.Render(sb);

            return sb.ToString();
        }

        private string GenerateFromPreset(RenderPresets preset)
        {
            switch (preset)
            {
                case RenderPresets.ContainerMargins:
                    return containerMargins();
                case RenderPresets.RotatingContainer:
                    return rotatingContainer();
            }
            return "";
        }

        //public string RenderTestCanvas(double widthPixel, double heightPixel, Color bgColor)
        //{

        //}

        //public string RenderBaseItemToSvg(DrawingBase drawing)
        //{

        //}

        ////Attempt at ensuring features can be created/tested individually by the system
        ////Without relying on having a whole workbook or epplus project.
        ////Simply: Does our positioning, sizing and parent hierarchy logic work as expected or not.
        //public string RenderIndependentCanvas(double widthPixel, double heightPixel, Color bgColor)
        //{
        //    var widthPoint = widthPixel.PixelToPoint();
        //    var heightPoint = heightPixel.PixelToPoint();

        //    var CanvasBounds = new BoundingBox(widthPoint, heightPoint);

        //    var sb = new StringBuilder();
        //    var canvas = new SvgIndependentCanvas(CanvasBounds, bgColor);

        //    var rect = new SvgIndependentRect(canvas.Bounds, widthPoint/2, heightPoint/2);
        //    rect.FillColor = "red";
        //    rect.Left = 10;

        //    //BoundingBox boundsTextBox = new BoundingBox(rect.Bounds.Left, rect.Bounds.Top, rect.Bounds.Width, rect.Bounds.Height);
        //    //var independentTxtBox = new SvgIndependentTextBox(canvas, boundsTextBox);

        //    //var textBox = new svgin

        //    canvas.AddRenderItem(rect);

        //    canvas.Render(sb);
        //    return sb.ToString();
        //}


        //public string RenderTextBox(ExcelDrawing someDrawing, BoundingBox parent, double maxHeight, double maxWidth)
        //{
        //    var sb = new StringBuilder();
        //    var svgTextBox = new SvgTextBox(someDrawing, parent, maxWidth, maxHeight);
        //}


        //public string RenderRangeToSvg(ExcelRange range)
        //{
        //    var ws = range.Worksheet;

        //    //ws.Workbook.Styles.CellXfs.

        //    double totalWidth = 0;
        //    foreach (var col in range.EntireColumn)
        //    {
        //        totalWidth += ExcelColumn.ColumnWidthToPixels(col.Width, range.Worksheet.Workbook.MaxFontWidth);
        //    }
        //    double totalHeight = 0;
        //    foreach (var row in range.EntireRow)
        //    {
        //        totalHeight = row.Height;
        //    }

        //    var sRange = new SvgRange(range, totalWidth, totalHeight);

        //    var sb = new StringBuilder();
        //    sb.Append($"<svg width=\"{500}\" height=\"{500}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\">");
        //    sRange.Render(sb);
        //    sb.Append($"</svg>");
        //    return sb.ToString();
        //}

        //public string RenderBox(string boxText)
        //{
        //    string retStr = "";
        //    var container = new TextContainerBase(boxText);
        //    var element = GenerateSvg(container);

        //    using (var ms = EPPlusMemoryManager.GetStream())
        //    {
        //        SvgWriter writer = new SvgWriter(ms, Encoding.UTF8);
        //        writer.RenderSvgElement(element, true);
        //        ms.Position = 0;
        //        using (var sr = new StreamReader(ms))
        //        {
        //            retStr = sr.ReadToEnd();
        //            return retStr;
        //        }
        //    }

        //    //writer.RenderSvgElement(element, true);

        //    //StreamReader reader = new StreamReader(ms);
        //    //retStr = reader.ReadToEnd();

        //    ////SvgParagraph para = new SvgParagraph(container.GetContent(),);
        //    ////var doc = new SvgEpplusDocument();
        //    //return retStr;
        //}

        //public string RenderTextBody(ExcelTextBody body, double shapeWidth, double shapeHeight)
        //{
        //    var sb = new StringBuilder();

        //    BoundingBox worldBounds = new BoundingBox();
        //    worldBounds.Width = shapeWidth;
        //    worldBounds.Height = shapeHeight;

        //    var doc = new SvgEpplusDocument((int)worldBounds.Width, (int)worldBounds.Height);
        //    doc.Render(sb);

        //    FontMeasurerTrueType measurer = new FontMeasurerTrueType(11, "Aptos Narrow", FontSubFamily.Regular);

        //    var svgBody = new SvgTextBodyItem(doc, worldBounds, null);
        //    svgBody.ImportTextBody(body);

        //    svgBody.Render(sb);
        //    sb.AppendLine("</svg>");

        //    return sb.ToString();
        //}

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

        //internal SvgElement GenerateSvg(TextContainerBase container)
        //{
        //    var fullString = container.GetContent();

        //    var doc = new SvgEpplusDocument(500, 500);

        //    var bg = new SvgElement("rect");
        //    bg.AddAttribute("width", "100%");
        //    bg.AddAttribute("height", "100%");
        //    bg.AddAttribute("fill", "red");
        //    bg.AddAttribute("opacity", "0.1");

        //    var nameId = "boundingBox";
        //    var def = new SvgElement("defs");
        //    var clipPath = new SvgElement("clipPath");
        //    clipPath.AddAttribute("id", nameId);

        //    def.AddChildElement(clipPath);

        //    var bb = new SvgElement("rect");
        //    bb.AddAttribute("x", container.Position.X);
        //    bb.AddAttribute("y", container.Position.Y);
        //    bb.AddAttribute("width", container.Width);
        //    bb.AddAttribute("height", container.Height);
        //    //bb.AddAttribute("fill", "blue");
        //    //bb.AddAttribute("opacity", "0.5");

        //    clipPath.AddChildElement(bb);

        //    var fontSizePx = 16d;

        //    var renderElement = new SvgElement("text");
        //    renderElement.AddAttribute("x", container.Position.X);
        //    renderElement.AddAttribute("y", container.Position.Y + fontSizePx);
        //    renderElement.AddAttribute("_measurementFont-size", $"{fontSizePx}px");
        //    renderElement.AddAttribute("clip-path", $"url(#{nameId})");

        //    renderElement.Content = fullString;

        //    var bbVisual = new SvgElement("rect");
        //    bbVisual.AddAttribute("x", container.Position.X);
        //    bbVisual.AddAttribute("y", container.Position.Y);
        //    bbVisual.AddAttribute("width", container.Width);
        //    bbVisual.AddAttribute("height", container.Height);
        //    bbVisual.AddAttribute("fill", "blue");
        //    bbVisual.AddAttribute("opacity", "0.5");

        //    doc.AddChildElement(def);
        //    doc.AddChildElement(bg);
        //    doc.AddChildElement(bbVisual);
        //    doc.AddChildElement(renderElement);

        //    doc.AddAttributes();

        //    return doc;
        //}
    }
}
