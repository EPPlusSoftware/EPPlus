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
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;

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

            txtBox.ImportTextBody(dataLabel.TextBody, false);

            txtBox.TextBody.Bounds.Top = 0;
            txtBox.TextBody.AutoSize = true;

            if (txtBox.TextBody.Paragraphs.Count == 0)
            {
                txtBox.TextBody.ImportParagraph(defaultParagraph, 0, finalString);
                //txtBox.TextBody.AddParagraph(0, finalString);
            }
            else if (txtBox.TextBody.Paragraphs.Count == 1)
            {
                txtBox.TextBody.ImportParagraph(dataLabel.TextBody.Paragraphs[0], 0, finalString);
                //Remove dummy paragraph added by ImportTextBody
                txtBox.TextBody.Paragraphs.RemoveAt(0);
            }

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
                    Rectangle = rect;

                    _manualLayoutOffset = new Coordinate(Rectangle.Left, Rectangle.Top);

                    Rectangle.Bounds.Left += _manualLayoutOffset.X;
                    Rectangle.Bounds.Top += _manualLayoutOffset.Y;

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
                    Rectangle.Bounds.Left -= _txtBox.Width + (point.Width / 2);
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    Rectangle.Bounds.Left += _txtBox.Width / 2 + point.Width;
                    break;
                case eLabelPosition.Top:
                    Rectangle.Bounds.Top -= (point.Height + _txtBox.Height) / 2;
                    break;
                case eLabelPosition.Bottom:
                    Rectangle.Bounds.Top += (point.Height + _txtBox.Height) / 2;
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position entered in SetPositionBasic: '{basicPosition}' is not a basic position");
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
                        dataLabelCenter = ((startToEndDir*0.5d) * -1d);
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
                case eLabelPosition.InEnd:
                    //if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.InEnd MUST have a direction." +
                    //        "Cannot be within End if EndPoint is undefined.");
                    //}
                    ////if(parentShape == null)
                    ////{
                    ////    throw new InvalidOperationException("eLabelPosition.InEnd MUST have a parentShape");
                    ////}
                    //startPointDirection = startPointDirection * -1;
                    //if (startPointDirection.X != 0)
                    //{
                    //    //If endPoint is to the right
                    //    if (startPointDirection.X > 0)
                    //    {
                    //        //We must place to the left
                    //        SetPositionBasic(parentPoint, eLabelPosition.Left);
                    //    }
                    //    //if endpoint is to the left
                    //    else
                    //    {
                    //        //We must place to the right
                    //        SetPositionBasic(parentPoint, eLabelPosition.Right);
                    //    }
                    //}

                    //if (startPointDirection.Y != 0)
                    //{
                    //    //If endpoint is below
                    //    if (startPointDirection.Y > 0)
                    //    {
                    //        //We must place on Top
                    //        SetPositionBasic(parentPoint, eLabelPosition.Top);
                    //    }
                    //    else
                    //    {
                    //        //We must place on Bottom
                    //        SetPositionBasic(parentPoint, eLabelPosition.Bottom);
                    //    }
                    //}
                    break;
                case eLabelPosition.OutEnd:
                    //if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    //{
                    //    throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a direction." +
                    //        "Cannot be within End if EndPoint is undefined.");
                    //}
                    ////if (parentShape == null)
                    ////{
                    ////    throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a parentShape");
                    ////}

                    //if (startPointDirection.X != 0)
                    //{
                    //    //If endPoint is to the left
                    //    if (startPointDirection.X < 0)
                    //    {
                    //        //We must place to the left
                    //        SetPositionBasic(parentPoint, eLabelPosition.Left);
                    //    }
                    //    //if endpoint is to the right
                    //    else
                    //    {
                    //        //We must place to the right
                    //        SetPositionBasic(parentPoint, eLabelPosition.Right);
                    //    }
                    //}

                    //if (startPointDirection.Y != 0)
                    //{
                    //    //If endpoint is on Top
                    //    if (startPointDirection.Y < 0)
                    //    {
                    //        //We must place on Top
                    //        SetPositionBasic(parentPoint, eLabelPosition.Top);
                    //    }
                    //    //If endpoint is on bottom
                    //    else
                    //    {
                    //        //We must place on Bottom
                    //        SetPositionBasic(parentPoint, eLabelPosition.Bottom);
                    //    }
                    //}
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
            renderItems.Add(parentPointGroup);

            var titleItemOrigin = new TitleRenderItem("DataLabel originpoint");
            renderItems.Add(titleItemOrigin);

            var group = new GroupRenderItem(Rectangle.Bounds);
            group.Left = Rectangle.Bounds.Left;
            group.Top = Rectangle.Bounds.Top;
            parentPointGroup.RenderItems.Add(group);

            var titleItem = new TitleRenderItem("DataLabel size adjustment");
            parentPointGroup.RenderItems.Add(titleItem);

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
                var iconGrp = new GroupRenderItem(new BoundingBox(_seriesIcon.Bounds.Left, height / 2 - 2));
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
