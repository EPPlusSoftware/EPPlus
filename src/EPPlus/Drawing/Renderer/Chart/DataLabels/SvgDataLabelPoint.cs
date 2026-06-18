using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.ShapeDefinitions;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Net;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgDataLabelPoint : ChartDrawingObject
    {
        bool _hasManualLayout = false;
        bool _hasLeaderLines = false;
        bool _haveAdjustedForIcon = false;
        bool _renderConnectionPointLines = false;

        private DrawingTextBox _txtBox;
        BoundingBox _parentPoint;
        List<LineRenderItem> _leaderLines = new List<LineRenderItem>();
        Coordinate _manualLayoutOffset = new Coordinate (0, 0);
        PointLines _connectionPointLines;
        eLabelPosition _labelPosition;

        //public SvgChartDataLabelStandard(DrawingChart chart, string dataLabelText) : base(chart)
        //{
        //    var txtBox = new SvgTextBox(chart, chart.Bounds, chart.Bounds);
        //    txtBox.AddText(0, dataLabelText);
        //}

        //public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard, SvgTextBox txtBox) : base(chart)
        //{
        //    HasLegendKey = standard.ShowLegendKey;
        //    TxtBox = txtBox;
        //}

        public SvgDataLabelPoint(ChartRenderer chart, ExcelChartDataLabelStandard standard) : base(chart)
        {
            _labelPosition = standard.Position;
            Rectangle = new RectRenderItem(chart.Bounds);
        }

        RenderItem _seriesIcon = null;

        internal void AddSeriesIcon(RenderItem seriesIcon)
        {
            var iconWidth = seriesIcon.Bounds.Width;
            var iconHeight = seriesIcon.Bounds.Height;

            _seriesIcon = seriesIcon;
            _seriesIcon.Bounds.Parent = Rectangle.Bounds;

            if (_haveAdjustedForIcon == false)
            {
                _txtBox.Left += iconWidth;
                _seriesIcon.Bounds.Left -= 0.75d;
                //It seems there is a hard-coded margin in excel of about 4.5pt (6px)
                Rectangle.Bounds.Left += 4d + 2.25d;
                LeftMargin -= 2.25d + 4d;
                Rectangle.Bounds.Width += iconWidth + 2.25d;

                _haveAdjustedForIcon = true;
            }
        }

        internal void ImportDataLabel(ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, ExcelDrawingParagraph defaultParagraph, BoundingBox maxBounds, BoundingBox defaultMargins)
        {
            List<string> dlblStrings = new List<string>();

            if (dataLabel.ShowSeriesName)
            {
                dlblStrings.Add(serie.GetHeaderString());
            }
            if (dataLabel.ShowCategory)
            {
                dlblStrings.Add(xValue.ToString());
            }
            if (dataLabel.ShowValue)
            {
                if (yValue != null)
                {
                    dlblStrings.Add(yValue.ToString());
                }
            }

            var separator = string.IsNullOrEmpty(dataLabel.Separator) ? ", " : dataLabel.Separator;

            string finalString = "";
            for (int j = 0; j < dlblStrings.Count; j++)
            {
                finalString += dlblStrings[j];
                if (j != dlblStrings.Count - 1)
                {
                    finalString += separator;
                }
            }

            var txtBox = new DrawingTextBox(Chart, Rectangle.Bounds, maxBounds.Width, maxBounds.Height);

            txtBox.ImportTextBodyAndParagraphs(dataLabel.TextBody, false);

            txtBox.TextBody.Bounds.Top = 0;
            txtBox.TextBody.AutoSize = true;

            if (txtBox.TextBody.Paragraphs.Count == 0)
            {
                if(defaultParagraph == null)
                {
                    txtBox.TextBody.AddParagraph(finalString);
                }
                else
                {
                    txtBox.ImportParagraph(defaultParagraph, 0, finalString);
                }
                //txtBox.TextBody.AddParagraph(0, finalString);
            }
            else if (txtBox.TextBody.Paragraphs.Count == 1)
            {
                txtBox.ImportParagraph(dataLabel.TextBody.Paragraphs[0], 0, finalString);
                //Remove dummy paragraph added by ImportTextBody
                txtBox.TextBody.Paragraphs.RemoveAt(0);
            }

            txtBox.TextBody.RecalculateParagraphs();

            if(txtBox.LeftMargin == 0)
            {
                txtBox.LeftMargin = defaultMargins.Left;
                txtBox.RightMargin = defaultMargins.Width;
                txtBox.TopMargin = defaultMargins.Top;
                txtBox.BottomMargin = defaultMargins.Height;
            }

            ////Txtbox is Broken. Workaround
            //txtBox.Left += txtBox.LeftMargin;
            //txtBox.Top += txtBox.TopMargin;

            //Center the textbox at the origin point
            Rectangle.Bounds.Left -= txtBox.Rectangle.Bounds.Width / 2;
            Rectangle.Bounds.Top -= txtBox.Rectangle.Bounds.Height / 2;

            //Set initial width and height to content
            Rectangle.Bounds.Width = txtBox.Rectangle.Bounds.Width;
            Rectangle.Bounds.Height = txtBox.Rectangle.Bounds.Height;

            _txtBox = txtBox;

            if (dataLabel.Fill.IsEmpty == false)
            {
                _txtBox.Rectangle.SetDrawingPropertiesFill(ChartRenderer.Theme, dataLabel.Fill, null);
            }
            //else
            //{
            //    _txtBox.Rectangle.SetDrawingPropertiesFill(ChartRenderer.Theme, dataLabel.Fill, null);
            //    _txtBox.Rectangle.FillColor = "transparent";
            //}

            if (dataLabel.Font.IsEmpty == false)
            {
                txtBox.TextBody.FontColorString = "#" + dataLabel.Font.Color.ToColorString();
            }

            _labelPosition = dataLabel.Position;

            if (dataLabel is ExcelChartDataLabelItem)
            {
                var individualLabel = dataLabel as ExcelChartDataLabelItem;

                if (individualLabel.Fill.IsEmpty == false)
                {
                    _txtBox.Rectangle.FillColor = "#" + individualLabel.Fill.Color.ToColorString();
                }

                if (individualLabel.Layout != null && individualLabel.Layout.HasLayout)
                {
                    _hasManualLayout = true;
                    var rect = GetRectFromManualLayout(ChartRenderer, individualLabel.Layout);

                    Rectangle.Bounds.Left += rect.Left;
                    Rectangle.Bounds.Top += rect.Top;

                    _manualLayoutOffset = new Coordinate(rect.Left, rect.Top);

                    if (rect.Bounds.Width != 0)
                    {
                        Rectangle.Bounds.Width = rect.Bounds.Width;
                    }
                    if (rect.Bounds.Height != 0)
                    {
                        Rectangle.Bounds.Height = rect.Bounds.Height;
                    }

                    if (dataLabel.ShowLeaderLines)
                    {
                        _hasLeaderLines = true;
                    }
                }
            }
        }

        private int GetClosestConnectionPointCoordinateIndex(Coordinate originPoint)
        {
            double smallestDist = double.MaxValue;
            int i = 0;
            int smallestIndex = 0;

            foreach (var line in _connectionPointLines.RenderLines)
            {
                line.X1 = originPoint.X;
                line.Y1 = originPoint.Y;

                var w = Math.Abs(line.X2 - line.X1);
                var h = Math.Abs(line.Y2 - line.Y1);

                //Use pythagoran theorem to get diagonal distance
                var totalDist = Math.Sqrt(Math.Pow(w, 2) + Math.Pow(h, 2));

                if (totalDist < smallestDist)
                {
                    smallestDist = totalDist;
                    smallestIndex = i;
                }
                i++;
            }

            return smallestIndex;
        }

        BoundingBox _parentShapeBounds = null;



        private void SetPositionBasic(BoundingBox point, eLabelPosition basicPosition)
        {
            switch (basicPosition)
            {
                //case eLabelPosition.Center:
                //    Bounds.Left = dataLabelCenter.X;
                //    Bounds.Top = dataLabelCenter.Y;
                //    break;
                case eLabelPosition.Left:
                    Rectangle.Bounds.Left -= (_txtBox.Width/2) + (point.Width / 2d);
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    Rectangle.Bounds.Left += (_txtBox.Width / 2d) + point.Width;
                    break;
                case eLabelPosition.Top:
                    Rectangle.Bounds.Top -= (point.Height + _txtBox.Height) / 2d;
                    break;
                case eLabelPosition.Bottom:
                    Rectangle.Bounds.Top += (point.Height + _txtBox.Height) / 2d;
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position entered in SetPositionBasic: '{basicPosition}' is not a basic position");
            }
        }


        RectRenderItem originPointRect;
        RectRenderItem basePositionRect;
        RectRenderItem endPositionRect;
        RectRenderItem centerPositionRect;


        private RectRenderItem GenerateDebugRenderItem(BoundingBox parent, string fillColor)
        {
            var pointRect = new RectRenderItem(parent);
            pointRect.Width = 10d;
            pointRect.Height = 10d;
            pointRect.FillColor = fillColor;
            pointRect.Left = -5d;
            pointRect.Top = -5d;
            return pointRect;
        }


        private void CreateDebugPoints(Transform basePoint, Transform endPoint, Transform centerPoint)
        {
            originPointRect = GenerateDebugRenderItem(_parentPoint, "darkRed");
            basePositionRect = GenerateDebugRenderItem(_parentPoint, "darkGreen");
            basePositionRect.Left += basePoint.LocalPosition.X;
            basePositionRect.Top += basePoint.LocalPosition.Y;
            endPositionRect = GenerateDebugRenderItem(_parentPoint, "darkBlue");
            endPositionRect.Left += endPoint.LocalPosition.X;
            endPositionRect.Top += endPoint.LocalPosition.Y;
            endPositionRect.BorderWidth = 2d;
            endPositionRect.BorderColor = "cyan";
            centerPositionRect = GenerateDebugRenderItem(_parentPoint, "Purple");
            centerPositionRect.Left += centerPoint.LocalPosition.X;
            centerPositionRect.Top += centerPoint.LocalPosition.Y;
        }

        /// <summary>
        /// 
        /// </summary>
        internal void SetShapeDimensions(Transform basePoint, Transform endPoint)
        {
            if(basePoint.Parent != endPoint.Parent)
            {
                throw new InvalidOperationException("basePoint and endPoint have different parents. " +
                    "Please ensure that they share the same parent");
            }


            //--- Set parent point ---
            _parentPoint = new BoundingBox();
            _parentPoint.Parent = basePoint.Parent.Parent;
            _parentPoint.Left = basePoint.Parent.LocalPosition.X;
            _parentPoint.Top = basePoint.Parent.LocalPosition.Y;
            Rectangle.Bounds.Parent = _parentPoint;
            //---

            //--- Calculate vectors and center point ---
            var endVector = endPoint.LocalPosition;
            var baseVector = basePoint.LocalPosition;

            var endToBaseVector = baseVector - endVector;
            var centerVector = endToBaseVector * 0.5d;

            //Translate from end point towards base point by 50% to find the center point
            Transform centerPoint = new Transform(endPoint.LocalPosition + centerVector, endPoint.LocalPosition + centerVector);
            centerPoint.Parent = basePoint.Parent;
            //--- 

            //At this point our rectangle globally is centered on the top-left of the object.
            //And endVector is the top center position.

            //--- Visualize positions for debugging purposes
            //CreateDebugPoints(basePoint, endPoint, centerPoint);
            //---

            switch (_labelPosition)
            {
                case eLabelPosition.Center:
                    Rectangle.Bounds.Left += centerPoint.LocalPosition.X;
                    Rectangle.Bounds.Top += centerPoint.LocalPosition.Y;
                    break;
                case eLabelPosition.Left:
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    break;
                case eLabelPosition.Top:
                    break;
                case eLabelPosition.Bottom:
                    break;
                case eLabelPosition.InBase:
                    //Translate to the base point
                    Rectangle.Bounds.Position += basePoint.LocalPosition;

                    //Move the textbox margins inside on left
                    if (endToBaseVector.X != 0)
                    {
                        //If basePoint is to the left
                        if (endToBaseVector.X > 0)
                        {
                            //We must place to the left
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Left);
                        }
                        //if basePoint is to the right
                        else
                        {
                            //We must place to the right
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Right);
                        }
                    }

                    //Move the textbox margins inside on top
                    if (endToBaseVector.Y != 0)
                    {
                        //If endpoint is on Top
                        if (endToBaseVector.Y > 0)
                        {
                            //We must place on bottom and apply margin to height
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Top);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on top and apply margin to height
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Bottom);
                        }
                    }
                    break;
                case eLabelPosition.InEnd:
                    //Move to end point
                    Rectangle.Bounds.Position += endPoint.LocalPosition;

                    //Move the textbox margins inside on left
                    if (endToBaseVector.X != 0)
                    {
                        //If basePoint is to the left
                        if (endToBaseVector.X < 0)
                        {
                            //We must place to the left
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Left);
                        }
                        //if basePoint is to the right
                        else
                        {
                            //We must place to the right
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Right);
                        }
                    }

                    //Move the textbox margins inside on top
                    if (endToBaseVector.Y != 0)
                    {
                        //If endpoint is on Top
                        if (endToBaseVector.Y < 0)
                        {
                            //We must place on bottom
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Top);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on top
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Bottom);
                        }
                    }
                    break;
                case eLabelPosition.OutEnd:
                    //Move to end point
                    Rectangle.Bounds.Position += endPoint.LocalPosition;
                    if (endToBaseVector.X != 0)
                    {
                        //If endPoint is to the left
                        if (endToBaseVector.X > 0)
                        {
                            //We must place to the left
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Left);
                        }
                        //if endpoint is to the right
                        else
                        {
                            //We must place to the right
                            SetPositionBasic(new BoundingBox(0, 0) { Width = 5d }, eLabelPosition.Right);
                        }
                    }

                    if (endToBaseVector.Y != 0)
                    {
                        //If endpoint is on Top
                        if (endToBaseVector.Y > 0)
                        {
                            //We must place on Top
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Top);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on Bottom
                            SetPositionBasic(new BoundingBox(0, 0) { Height = 5d }, eLabelPosition.Bottom);
                        }
                    }

                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position {_labelPosition} has not been implemented yet");
            }
        }

        internal void SetParentPoint(BoundingBox parentPoint, BoundingBox parentShape, Vector2 startToEndDir)
        {
            Rectangle.Bounds.Parent = parentPoint;
            _parentPoint = parentPoint;
            _parentShapeBounds = parentShape;

            var dataLabelCenter = new Vector2(Rectangle.Bounds.Left, Rectangle.Bounds.Top);
            Vector2 startPointDirection = Vector2.Zero;

            if ((startToEndDir.X == 0 && startToEndDir.Y == 0) == false)
            {
                startPointDirection = startToEndDir / startToEndDir.Length;
            }

            //Rectangle.Bounds.Left += 20;

            //if (_parentShapeBounds != null)
            //{
            //    //shapeCenter
            //    dataLabelCenter = new Graphics.Math.Vector2((_parentShapeBounds.Width / 2)+parentShape.Left, (_parentShapeBounds.Height / 2)+parentShape.Top);

            //    //Get directional vector (in local coords but does not matter since we make it directional)
            //    startPointDirection = dataLabelCenter - parentPoint.LocalPosition;
            //    //Divide by length to only get direction
            //    var startPointDirectionOnly = startPointDirection / startPointDirection.Length;

            //    //var lenX = Math.Abs()
            //    //////Pythagoran theorem
            //    var len = Math.Sqrt(Math.Pow(startPointDirection.X, 2) + Math.Pow(startPointDirection.Y, 2));
            //    startPointDirection = startPointDirectionOnly * len;
            //}
            //else
            //{
            //    dataLabelCenter = new Graphics.Math.Vector2(Bounds.Left, Bounds.Top);
            //}

            switch (_labelPosition)
            {
                case eLabelPosition.Center:

                    if ((startToEndDir.X == 0 && startToEndDir.Y == 0) == false)
                    {
                        //Half and invert
                        dataLabelCenter = ((startToEndDir * 0.5d) * -1d);
                    }
                    else if (startToEndDir.Y != 0)
                    {
                        //Half and invert
                        dataLabelCenter = ((startToEndDir * 0.5d) * -1d);
                    }

                    Rectangle.Bounds.Left += dataLabelCenter.X;
                    Rectangle.Bounds.Top += dataLabelCenter.Y;
                    break;
                case eLabelPosition.Left:
                    SetPositionBasic(parentPoint, _labelPosition);
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    SetPositionBasic(parentPoint, _labelPosition);
                    break;
                case eLabelPosition.Top:
                    SetPositionBasic(parentPoint, _labelPosition);
                    break;
                case eLabelPosition.Bottom:
                    SetPositionBasic(parentPoint, _labelPosition);
                    break;
                case eLabelPosition.InBase:
                    var endToStartVector = startToEndDir * -1;
                    //Rectangle.Left *= endToStartVector.X;
                    //Rectangle.Top *= endToStartVector.Y;
                    if (startPointDirection.X != 0)
                    {
                        Rectangle.Left += endToStartVector.X;
                        //If basePoint is to the left
                        if (startPointDirection.X < 0)
                        {
                            //We must place to the left
                            SetPositionBasic(new BoundingBox(0,0), eLabelPosition.Left);
                        }
                        //if basePoint is to the right
                        else
                        {
                            //We must place to the right
                            SetPositionBasic(new BoundingBox(0, 0), eLabelPosition.Right);
                        }
                    }

                    if (startPointDirection.Y != 0)
                    {
                        Rectangle.Top += endToStartVector.Y;
                        //If endpoint is on Top
                        if (startPointDirection.Y < 0)
                        {
                            //We must place on bottom
                            SetPositionBasic(new BoundingBox(0, 0), eLabelPosition.Top);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on top
                            SetPositionBasic(new BoundingBox(0, 0), eLabelPosition.Bottom);
                        }
                    }
                    break;
                case eLabelPosition.InEnd:
                    if (startPointDirection.X != 0)
                    {
                        //If endPoint is to the left
                        if (startPointDirection.X < 0)
                        {
                            //We must place to the right
                            SetPositionBasic(parentPoint, eLabelPosition.Right);
                        }
                        //if endpoint is to the right
                        else
                        {
                            //We must place to the left
                            SetPositionBasic(parentPoint, eLabelPosition.Left);
                        }
                    }

                    if (startPointDirection.Y != 0)
                    {
                        //If endpoint is on Top
                        if (startPointDirection.Y < 0)
                        {
                            //We must place on bottom
                            SetPositionBasic(parentPoint, eLabelPosition.Bottom);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on top
                            SetPositionBasic(parentPoint, eLabelPosition.Top);
                        }
                    }
                    //if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.InEnd MUST have a direction." +
                    //        "Cannot be within End if EndPoint is undefined.");
                    //}
                    //var insidePos = startToEndDir * 0.15 * -1;
                    //Rectangle.Bounds.Left += insidePos.X;
                    //Rectangle.Bounds.Top += insidePos.Y;
                    break;
                case eLabelPosition.OutEnd:
                    //if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a direction." +
                    //        "Cannot be within End if EndPoint is undefined.");
                    //}
                    //if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a direction." +
                    //        "Cannot be within End if EndPoint is undefined.");
                    //}
                    //Rectangle.Bounds.Left += startToEndDir.X * 0.15;
                    //Rectangle.Bounds.Top += startToEndDir.Y * 0.15;
                    //if (parentShape == null)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a parentShape");
                    //}

                    if (startPointDirection.X != 0)
                    {
                        //If endPoint is to the left
                        if (startPointDirection.X < 0)
                        {
                            //We must place to the left
                            SetPositionBasic(parentPoint, eLabelPosition.Left);
                        }
                        //if endpoint is to the right
                        else
                        {
                            //We must place to the right
                            SetPositionBasic(parentPoint, eLabelPosition.Right);
                        }
                    }

                    if (startPointDirection.Y != 0)
                    {
                        //If endpoint is on Top
                        if (startPointDirection.Y < 0)
                        {
                            //We must place on Top
                            SetPositionBasic(parentPoint, eLabelPosition.Top);
                        }
                        //If endpoint is on bottom
                        else
                        {
                            //We must place on Bottom
                            SetPositionBasic(parentPoint, eLabelPosition.Bottom);
                        }
                    }
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position {_labelPosition} has not been implemented yet");
            }

            if (_hasManualLayout)
            {
                if (_hasLeaderLines)
                {
                    //With origin in top left of current bounds get the connection points
                    var cPoints = new ConnectionPointsMiddle(0, 0, Rectangle.Bounds.Width, Rectangle.Bounds.Height);

                    //Ready to draw the lines so that we can visualize the distances to each point
                    _connectionPointLines = new PointLines(ChartRenderer, Rectangle.Bounds, cPoints);

                    //Adjust if there is a margin
                    _connectionPointLines.Rectangle.Bounds.Left += LeftMargin;
                    _connectionPointLines.UpdateLines();

                    //Get the offset between those points and the origin point
                    var offsetToParentPoint = new Coordinate(-(Rectangle.Bounds.Left + LeftMargin), -(Rectangle.Bounds.Top + TopMargin));

                    //Calculate closest point
                    var index = GetClosestConnectionPointCoordinateIndex(offsetToParentPoint);

                    _leaderLines.Clear();

                    double xOffset = 0;
                    if (index == 0 || index == 2)
                    {
                        //If Left or Right
                        //Add extra 7 px (5.25pt) line to the given side
                        var extraLine = new LineRenderItem(ChartRenderer.Bounds);

                        xOffset += index == 0 ? -5.25d : 5.25d;

                        extraLine.X1 = _connectionPointLines.ConnectionPoints.Points[index].X + LeftMargin;
                        extraLine.Y1 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.Y2 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.X2 = extraLine.X1 + xOffset;

                        extraLine.BorderColor = "gray";
                        extraLine.BorderWidth = 0.5;

                        _leaderLines.Add(extraLine);
                    }
                    var mainLine = new LineRenderItem(ChartRenderer.Bounds);
                    mainLine.X1 = _connectionPointLines.ConnectionPoints.Points[index].X + xOffset + LeftMargin;
                    mainLine.Y1 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                    mainLine.X2 = offsetToParentPoint.X + LeftMargin;
                    mainLine.Y2 = offsetToParentPoint.Y;

                    mainLine.BorderColor = "gray";
                    mainLine.BorderWidth = 0.5;
                    _leaderLines.Add(mainLine);
                }
            }
        }

        private void AppendDebugBounds(List<RenderItem> renderItems)
        {
            var rect = new RectRenderItem(Rectangle.Bounds);
            rect.Bounds.Left = LeftMargin;
            rect.Bounds.Top = 0;
            rect.Bounds.Width = Rectangle.Bounds.Width;
            rect.Bounds.Height = Rectangle.Bounds.Height;

            rect.FillColor = "red";
            rect.FillOpacity = 0.2;
            renderItems.Add(rect);
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var parentPointGroup = new GroupRenderItem(_parentPoint);
            parentPointGroup.Left = _parentPoint.Left;
            parentPointGroup.Top = _parentPoint.Top;

            var titleItemOrigin = new TitleRenderItem("DataLabel originpoint");
            parentPointGroup.AddChildItem(titleItemOrigin);

            if(originPointRect != null)
            {
                parentPointGroup.AddChildItem(originPointRect);
            }
            if(basePositionRect != null)
            {
                parentPointGroup.AddChildItem(basePositionRect);
            }
            if(endPositionRect != null)
            {
                parentPointGroup.AddChildItem(endPositionRect);
            }
            if(centerPositionRect != null)
            {
                parentPointGroup.AddChildItem(centerPositionRect);
            }

            renderItems.Add(parentPointGroup);

            var group = new GroupRenderItem(Rectangle.Bounds);
            group.Left = Rectangle.Bounds.Left;
            group.Top = Rectangle.Bounds.Top;

            var titleItem = new TitleRenderItem("DataLabel size adjustment");
            group.AddChildItem(titleItem);

            parentPointGroup.RenderItems.Add(group);

            _txtBox.AppendRenderItems(group.RenderItems);
            
            if(_renderConnectionPointLines)
            {
                if (_connectionPointLines != null)
                {
                    _connectionPointLines.AppendRenderItems(group.RenderItems);
                }
            }

            if (_seriesIcon != null)
            {
                var height = Rectangle.Bounds.Height;
                if (height == 0)
                {
                    height = _txtBox.Height;
                }
                //Currently series icon always has a y1 y2 of 2
                var iconGrp = new GroupRenderItem(new BoundingBox(_seriesIcon.Bounds.Left, height / 2));
                iconGrp.Left = _seriesIcon.Bounds.Left;
                iconGrp.Top = (height / 2) - 2;
                group.RenderItems.Add(iconGrp);
                iconGrp.RenderItems.Add(_seriesIcon);
            }

            if (_leaderLines != null && _leaderLines.Count > 0)
            {
                foreach (var line in _leaderLines)
                {
                    group.RenderItems.Add(line);
                }
            }
        }
    }
}
