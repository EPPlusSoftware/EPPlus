using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Export.ImageRenderer.Svg.DefinitionUtils;
using EPPlus.Export.ImageRenderer.Svg.DefinitionUtils.UtillNodes;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Export.ImageRenderer.Svg.Writer;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderTool
{
    internal class RenderItemClassGenerator
    {
        DrawingItemForTesting SvgCanvas;

        internal RenderItemClassGenerator(double width, double height)
        {
            SvgCanvas = new DrawingItemForTesting(new BoundingBox(width.PixelToPoint(), height.PixelToPoint()));

        }

        public enum RenderItemClasses
        {
            Rect,
            TextBox,
            Shape,
            CircleSegment
        }

        SvgTextBox GenerateTextBox(DrawingBase baseRenderer, BoundingBox parent, BoundingBox maxBounds)
        {
            return new SvgTextBox(baseRenderer, parent, maxBounds);
        }


        public Dictionary<string, object> GetItemProperties(RenderItemClasses item)
        {
            switch (item)
            {
                case RenderItemClasses.Rect:
                    return new Dictionary<string, object> { { "Top", 10d }, { "Left", 10d }, { "Width", 10d }, { "Height", 10d }, { "Opacity", 0.8 }, { "Fill", Color.Goldenrod } };
                case RenderItemClasses.CircleSegment:
                    return new Dictionary<string, object> { { "angle", 90d }, { "radius", 144d }, { "cx", 144d }, { "cy", 144d } };
                case RenderItemClasses.TextBox:
                default:
                    throw new NotImplementedException("This class has not been implemented as an option yet");

            }
        }

        private string RenderCircleSegment(DrawingBase baseItem, Dictionary<string, object> itemProperties)
        {
            return RenderCircleSegment((double)itemProperties["angle"], (double)itemProperties["radius"], (double)itemProperties["cx"], (double)itemProperties["cy"]);
        }

        string RenderCircleSegment(double degree, double radius, double cx, double cy)
        {
            degree = degree % 360;

            if (degree < 0)
            {
                degree = 360 - degree;
            }

            //Adjust by -90 so it starts from the top
            var angleRadians = (degree - 90d) * (Math.PI / 180.0d);

            //radius = radius.PixelToPoint();
            //cx = radius.PixelToPoint();
            //cy = radius.PixelToPoint();

            var xPoint = cx + (radius * Math.Cos(angleRadians));
            var yPoint = cy + (radius * Math.Sin(angleRadians));

            Coordinate endPoint = new Coordinate(xPoint, yPoint);

            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72 * 4;
            baseBB.Height = 72 * 4;

            var baseItem = new DrawingItemForTesting(baseBB);

            BoundingBox parent = new BoundingBox();

            //Transform rotationPoint = new Transform();
            //rotationPoint.SetLocalPositionWithWorldCoordinates(new Vector2(cx, cy));
            //var item = new SvgGroupItemNew(baseItem, parent, -45d, rotationPoint);

            var slice = new SvgRenderPathItem(baseItem, baseItem.Bounds);

            //item.AddChildItem(slice);

            var startPoint = new Coordinate(cx, cy - radius);

            var moveCenter = new PathCommands(PathCommandType.Move, slice, cx / baseItem.Bounds.Width, cy / baseItem.Bounds.Height);
            var lineToStart = new PathCommands(PathCommandType.Line, slice, startPoint.X / baseItem.Bounds.Width, startPoint.Y / baseItem.Bounds.Height);

            var w = baseItem.Bounds.Width.PointToPixel();
            var h = baseItem.Bounds.Height.PointToPixel();

            var radX = radius.PointToPixel() / w;
            var radY = radius.PointToPixel() / h;

            var arcCommand = new PathCommands(PathCommandType.Arc, slice, new double[] { radX, radY, 0, degree > 180 ? 1 : 0, 1, endPoint.X / baseItem.Bounds.Width, endPoint.Y / baseItem.Bounds.Height });

            slice.Commands.Add(moveCenter);
            slice.Commands.Add(lineToStart);
            slice.Commands.Add(arcCommand);

            slice.FillColor = "red";
            slice.BorderColor = "green";

            baseItem.RenderItems.Add(slice);

            var sb = new StringBuilder();

            baseItem.Render(sb);

            return sb.ToString();

            //return baseItem;
        }

        public SvgRenderRectItem AddRect(SvgGroupItemNew parentGroup = null)
        {
            SvgRenderRectItem rectItem;

            if (parentGroup == null)
            {
                rectItem = new SvgRenderRectItem(SvgCanvas, SvgCanvas.Bounds);
                SvgCanvas.RenderItems.Add(rectItem);
            }
            else
            {
                rectItem = new SvgRenderRectItem(SvgCanvas, parentGroup.Bounds);
                parentGroup.AddChildItem(rectItem);
            }
            return rectItem;
        }

        //SvgRenderRectItem GenerateRect(DrawingBase baseItem)
        //{
        //    var rect = new SvgRenderRectItem(baseItem, baseItem.Bounds);

        //    return rect;
        //}

        //public RenderItem GenerateItem(RenderItemClasses preset, DrawingBase baseItem)
        //{
 
        //    switch (preset)
        //    {
        //        case RenderItemClasses.Rect:
        //            return GenerateRect(baseItem);


        //        default:
        //            throw new NotImplementedException("This class has not been implemented as an option yet");
        //    }
        //}

        //private RenderItem GenerateClass(RenderItemClasses preset, DrawingBase baseItem)
        //{
        //    switch (preset)
        //    {
        //        case RenderItemClasses.Rect:
        //            return GenerateRect(baseItem);


        //        default:
        //            throw new NotImplementedException("This class has not been implemented as an option yet");
        //    }
        //}
        /////// <summary>
        /////// For Testing the specific renderItem class
        /////// </summary>
        /////// <param name="item"></param>
        /////// <param name="width"></param>
        /////// <param name="height"></param>
        /////// <returns></returns>
        ////public RenderItem AddIndividualClass(RenderItemClasses item)
        ////{
        ////    GenerateFromClasses
        ////    var svgCanvas = new DrawingItemForTesting(new BoundingBox(width.PixelToPoint(), height.PixelToPoint()));

        ////    RenderItem renderItem;

        ////    if (item == RenderItemClasses.CircleSegment)
        ////    {
        ////        return GenerateFromCircle(item, svgCanvas, itemProperties);
        ////    }
        ////    else
        ////    {
        ////        renderItem = GenerateFromClasses(item, svgCanvas, itemProperties);
        ////    }

        ////    svgCanvas.RenderItems.Add(renderItem);

        ////    var sb = new StringBuilder();

        ////    svgCanvas.Render(sb);

        ////    return sb.ToString();
        ////}

        //public enum RenderPresets
        //{
        //    ContainerMargins,
        //    RotatingContainer,
        //    PatternFill,
        //}

        //public string RenderTest(RenderPresets preset)
        //{
        //    return GenerateFromPreset(preset);
        //}


        //private string rotatingContainer()
        //{
        //    var baseBB = new BoundingBox();

        //    //96x96 px
        //    baseBB.Width = 72;
        //    baseBB.Height = 72;

        //    var baseItem = new DrawingItemForTesting(baseBB);

        //    BoundingBox parent = new BoundingBox();

        //    var groupItem = new SvgGroupItemNew(baseItem, parent, 45);

        //    groupItem.Position.Left = 10;
        //    groupItem.Position.Top = 10;

        //    SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

        //    rectItem.FillColor = "red";
        //    rectItem.FillOpacity = 0.2d;

        //    rectItem.Width = 20;
        //    rectItem.Height = 20;

        //    groupItem.AddChildItem(rectItem);


        //    SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
        //    siblingItem.FillColor = "blue";
        //    siblingItem.FillOpacity = 0.2d;

        //    siblingItem.Width = 20;
        //    siblingItem.Height = 20;

        //    siblingItem.Bounds.Left = 20;
        //    siblingItem.Bounds.Top = 20;

        //    groupItem.AddChildItem(siblingItem);

        //    groupItem.SetRotationPointToCenterOfGroup();

        //    SvgRenderRectItem centerOfGroupMarker = new SvgRenderRectItem(baseItem, baseItem.Bounds);
        //    centerOfGroupMarker.FillColor = "green";
        //    centerOfGroupMarker.FillOpacity = 0.8d;

        //    centerOfGroupMarker.Width = 6;
        //    centerOfGroupMarker.Height = 6;

        //    centerOfGroupMarker.Left = 30 - (centerOfGroupMarker.Width / 2);
        //    centerOfGroupMarker.Top = 30 - (centerOfGroupMarker.Height / 2);

        //    baseItem.RenderItems.Add(centerOfGroupMarker);

        //    var sb = new StringBuilder();

        //    baseItem.RenderItems.Add(groupItem);

        //    baseItem.Render(sb);

        //    return sb.ToString();
        //}

        //private string containerMargins()
        //{
        //    var baseBB = new BoundingBox();

        //    baseBB.Width = 400;
        //    baseBB.Height = 400;

        //    var baseItem = new DrawingItemForTesting(baseBB);

        //    SvgRenderRectItem myBgItem = new SvgRenderRectItem(baseItem, baseItem.Bounds);
        //    myBgItem.FillColor = "purple";
        //    myBgItem.FillOpacity = 0.2d;

        //    SvgRenderRectItem myInnerItem = new SvgRenderRectItem(baseItem, myBgItem.Bounds);

        //    myInnerItem.FillColor = "green";
        //    myInnerItem.FillOpacity = 0.8d;

        //    myInnerItem.Width = 50;
        //    myInnerItem.Height = 50;

        //    var container = new SvgContainerItem(myInnerItem, myBgItem);

        //    container.MarginLeft = 5;
        //    container.MarginRight = 5;
        //    container.MarginTop = 5;
        //    container.MarginBottom = 5;

        //    container.ApplyMargins();

        //    baseItem.RenderItems.Add(container);

        //    var sb = new StringBuilder();

        //    baseItem.Render(sb);

        //    return sb.ToString();
        //}

        //internal string pattern()
        //{
        //    var baseBB = new BoundingBox();

        //    baseBB.Width = 400;
        //    baseBB.Height = 400;

        //    var baseItem = new DrawingItemForTesting(baseBB);


        //    var linePattern = new LinePattern(baseItem, "testLines", LinePatternType.Vertical);
        //    linePattern.SetNumberOfLines(3);

        //    var defItem = new DefinitionGroup(baseItem);
        //    defItem.Items.Add(linePattern);

        //    var rectItem = new SvgRenderRectItem(baseItem, baseItem.Bounds);

        //    rectItem.FillColor = $"url(#{"testLines"})";

        //    rectItem.Width = 200;
        //    rectItem.Height = 200;

        //    baseItem.RenderItems.Add(defItem);
        //    baseItem.RenderItems.Add(rectItem);
        //    var useItem = new SvgUseRefItem(baseItem, baseItem.Bounds, "testLines");

        //    baseItem.RenderItems.Add(useItem);

        //    var sb = new StringBuilder();

        //    baseItem.Render(sb);

        //    return sb.ToString();
        //}

        //internal string pattern2()
        //{
        //    var baseBB = new BoundingBox();

        //    baseBB.Width = 400;
        //    baseBB.Height = 400;

        //    var baseItem = new DrawingItemForTesting(baseBB);

        //    string refId = "grid";

        //    var defItem = new DefinitionGroup(baseItem);

        //    var dynaGrid = new DynamicGridDefGroup(baseItem, refId, 7, 5);
        //    defItem.Items.Add(dynaGrid);

        //    baseItem.RenderItems.Add(defItem);

        //    var useItem = new SvgUseRefItem(baseItem, baseItem.Bounds, refId);
        //    useItem.Bounds.Width = 300;
        //    useItem.Bounds.Height = 200;

        //    baseItem.RenderItems.Add(useItem);

        //    var sb = new StringBuilder();

        //    baseItem.Render(sb);

        //    return sb.ToString();
        //}

        //private string GenerateFromPreset(RenderPresets preset)
        //{
        //    switch (preset)
        //    {
        //        case RenderPresets.ContainerMargins:
        //            return containerMargins();
        //        case RenderPresets.RotatingContainer:
        //            return rotatingContainer();
        //        case RenderPresets.PatternFill:
        //            return pattern2();
        //    }
        //    return "";
        //}

        ////private RenderItem GenerateRect(DrawingBase baseItem, Dictionary<string, object> itemProperties)
        ////{
        ////    var rectItem = new SvgRenderRectItem(baseItem, baseItem.Bounds);

        ////    if (itemProperties.ContainsKey("Top"))
        ////    {
        ////        rectItem.Top = (double)itemProperties["Top"];
        ////    }
        ////    if (itemProperties.ContainsKey("Left"))
        ////    {
        ////        rectItem.Left = (double)itemProperties["Left"];
        ////    }
        ////    if (itemProperties.ContainsKey("Width"))
        ////    {
        ////        rectItem.Width = (double)itemProperties["Width"];
        ////    }
        ////    if (itemProperties.ContainsKey("Height"))
        ////    {
        ////        rectItem.Height = (double)itemProperties["Height"];
        ////    }
        ////    if (itemProperties.ContainsKey("Opacity"))
        ////    {
        ////        rectItem.FillOpacity = (double)itemProperties["Opacity"];
        ////    }
        ////    if (itemProperties.ContainsKey("Fill"))
        ////    {
        ////        rectItem.FillColor = "#" + ((Color)itemProperties["Fill"]).ToColorString();
        ////    }

        ////    return rectItem;
        ////}

        ////private string GenerateFromCircle(RenderItemClasses preset, DrawingBase baseItem, Dictionary<string, object> itemProperties)
        ////{
        ////    return RenderCircleSegment(baseItem, itemProperties);
        ////}

        ////private RenderItem GenerateFromClasses(RenderItemClasses preset, DrawingBase baseItem, Dictionary<string, object> itemProperties)
        ////{
        ////    switch (preset)
        ////    {
        ////        case RenderItemClasses.Rect:
        ////            return GenerateRect(baseItem, itemProperties);
        ////        case RenderItemClasses.CircleSegment:
        ////        //return RenderCircleSegment(baseItem, itemProperties);
        ////        case RenderItemClasses.TextBox:


        ////        default:
        ////            throw new NotImplementedException("This class has not been implemented as an option yet");
        ////    }
        ////}

        ////internal string RenderSvgElement(SvgElement element)
        ////{
        ////    string retStr = string.Empty;

        ////    using (var ms = EPPlusMemoryManager.GetStream())
        ////    {
        ////        SvgWriter writer = new SvgWriter(ms, Encoding.UTF8);
        ////        writer.RenderSvgElement(element, true);
        ////        ms.Position = 0;
        ////        using (var sr = new StreamReader(ms))
        ////        {
        ////            retStr = sr.ReadToEnd();
        ////            return retStr;
        ////        }
        ////    }
        ////}

        ////internal SvgElement GetDefinitions(BoundingBox boundingBox, out string nameId, bool AllowOverflow = false)
        ////{
        ////    nameId = "boundingBox";
        ////    var def = new SvgElement("defs");

        ////    string defaultName = "defaultRect";

        ////    if (AllowOverflow == false)
        ////    {
        ////        var bb = new SvgElement("rect");
        ////        bb.AddAttribute("width", boundingBox.Width);
        ////        bb.AddAttribute("height", boundingBox.Height);
        ////        bb.AddAttribute("id", defaultName);

        ////        def.AddChildElement(bb);

        ////        var clipPath = new SvgElement("clipPath");
        ////        clipPath.AddAttribute("id", nameId);

        ////        def.AddChildElement(clipPath);

        ////        var useElement = new SvgElement("use");
        ////        useElement.AddAttribute("href", $"#{defaultName}");

        ////        clipPath.AddChildElement(useElement);
        ////    }

        ////    return def;
        ////}

    }
}
