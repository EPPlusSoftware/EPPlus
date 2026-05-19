using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgDataLabelPoint : SvgChartObject
    {
        bool _hasManualLayout = false;
        bool _hasLeaderLines = false;
        bool _haveAdjustedForIcon = false;
        bool _renderConnectionPointLines = false;

        private SvgTextBox _txtBox;
        BoundingBox _parentPoint;
        List<SvgRenderLineItem> _leaderLines = new List<SvgRenderLineItem>();
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

        public SvgDataLabelPoint(DrawingChart chart, ExcelChartDataLabelStandard standard) : base(chart)
        {
            _labelPosition = standard.Position;
        }

        RenderItem _seriesIcon = null;

        internal void AddSeriesIcon(RenderItem seriesIcon)
        {
            var iconWidth = seriesIcon.Bounds.Width;
            var iconHeight = seriesIcon.Bounds.Height;

            _seriesIcon = seriesIcon;
            _seriesIcon.Bounds.Parent = Bounds;

            if (_haveAdjustedForIcon == false)
            {
                _txtBox.Left += iconWidth;
                _seriesIcon.Bounds.Left -= 0.75d;
                //It seems there is a hard-coded margin in excel of about 4.5pt (6px)
                Bounds.Left += 4d + 2.25d;
                LeftMargin -= 2.25d + 4d;
                Bounds.Width += iconWidth + 2.25d;

                _haveAdjustedForIcon = true;
            }
        }

        internal void ImportDataLabel(SvgChart chart, ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, ExcelDrawingParagraph defaultParagraph, BoundingBox maxBounds, BoundingBox defaultMargins)
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

            var txtBox = new SvgTextBox(chart, Bounds, maxBounds);

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

            //Center the textbox at the origin point
            Bounds.Left -= txtBox.Rectangle.Bounds.Width / 2;
            Bounds.Top -= txtBox.Rectangle.Bounds.Height / 2;

            //Set initial width and height to content
            Bounds.Width = txtBox.Rectangle.Bounds.Width;
            Bounds.Height = txtBox.Rectangle.Bounds.Height;

            _txtBox = txtBox;

            if (dataLabel.Fill.IsEmpty == false)
            {
                _txtBox.Rectangle.SetDrawingPropertiesFill(dataLabel.Fill, null);
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
                    var rect = GetRectFromManualLayout(chart, individualLabel.Layout);
                    Rectangle = rect;

                    _manualLayoutOffset = new Coordinate(Rectangle.Left, Rectangle.Top);

                    Bounds.Left += _manualLayoutOffset.X;
                    Bounds.Top += _manualLayoutOffset.Y;

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
                    Bounds.Left -= _txtBox.Width + (point.Width / 2);
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    Bounds.Left += _txtBox.Width / 2 + point.Width;
                    break;
                case eLabelPosition.Top:
                    Bounds.Top -= (point.Height + _txtBox.Height) / 2;
                    break;
                case eLabelPosition.Bottom:
                    Bounds.Top += (point.Height + _txtBox.Height) / 2;
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position entered in SetPositionBasic: '{basicPosition}' is not a basic position");
            }
        }

        internal void SetParentPoint(BoundingBox parentPoint, Graphics.Math.Vector2 startToEndDir)
        {
            Bounds.Parent = parentPoint;
            _parentPoint = parentPoint;

            var dataLabelCenter = new Graphics.Math.Vector2(Bounds.Left, Bounds.Top);
            Graphics.Math.Vector2 startPointDirection = Graphics.Math.Vector2.Zero;

            if ((startToEndDir.X == 0 && startToEndDir.Y == 0) == false)
            {
                startPointDirection = startToEndDir / startToEndDir.Length;
            }

            switch (_labelPosition)
            {
                case eLabelPosition.Center:

                    if ((startToEndDir.X == 0 && startToEndDir.Y == 0) == false)
                    {
                        //Half and invert
                        dataLabelCenter = ((startToEndDir * 0.5d) * -1d);
                    }
                    Bounds.Left += dataLabelCenter.X;
                    Bounds.Top += dataLabelCenter.Y;
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
                    if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    {
                        throw new InvalidOperationException("eLabelPosition.InEnd MUST have a direction." +
                            "Cannot be within End if EndPoint is undefined.");
                    }
                    var insidePos = startToEndDir * 0.15 * -1;
                    Bounds.Left += insidePos.X;
                    Bounds.Top += insidePos.Y;
                    break;
                case eLabelPosition.OutEnd:
                    if (startPointDirection.X == 0 && startPointDirection.Y == 0)
                    {
                        throw new InvalidOperationException("eLabelPosition.OutEnd MUST have a direction." +
                            "Cannot be within End if EndPoint is undefined.");
                    }
                    Bounds.Left += startToEndDir.X * 0.15;
                    Bounds.Top += startToEndDir.Y * 0.15;
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position {_labelPosition} has not been implemented yet");
            }

            if (_hasManualLayout)
            {
                if (_hasLeaderLines)
                {
                    //With origin in top left of current bounds get the connection points
                    var cPoints = new ConnectionPointsMiddle(0, 0, Bounds.Width, Bounds.Height);

                    //Ready to draw the lines so that we can visualize the distances to each point
                    _connectionPointLines = new PointLines(ChartRenderer, Bounds, cPoints);

                    //Adjust if there is a margin
                    _connectionPointLines.Bounds.Left += LeftMargin;
                    _connectionPointLines.UpdateLines();

                    //Get the offset between those points and the origin point
                    var offsetToParentPoint = new Coordinate(-(Bounds.Left + LeftMargin), -(Bounds.Top + TopMargin));

                    //Calculate closest point
                    var index = GetClosestConnectionPointCoordinateIndex(offsetToParentPoint);

                    _leaderLines.Clear();

                    double xOffset = 0;
                    if (index == 0 || index == 2)
                    {
                        //If Left or Right
                        //Add extra 7 px (5.25pt) line to the given side
                        var extraLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);

                        xOffset += index == 0 ? -5.25d : 5.25d;

                        extraLine.X1 = _connectionPointLines.ConnectionPoints.Points[index].X + LeftMargin;
                        extraLine.Y1 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.Y2 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.X2 = extraLine.X1 + xOffset;

                        extraLine.BorderColor = "gray";
                        extraLine.BorderWidth = 0.5;

                        _leaderLines.Add(extraLine);
                    }
                    var mainLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);
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
            SvgRenderRectItem rect = new SvgRenderRectItem(ChartRenderer, Bounds);
            rect.Bounds.Left = LeftMargin;
            rect.Bounds.Top = 0;
            rect.Bounds.Width = Bounds.Width;
            rect.Bounds.Height = Bounds.Height;

            rect.FillColor = "red";
            rect.FillOpacity = 0.2;
            renderItems.Add(rect);
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var parentPointGroup = new SvgGroupItem(ChartRenderer, _parentPoint);
            renderItems.Add(parentPointGroup);

            var titleItemOrigin = new SvgTitleItem(DrawingRenderer, "DataLabel originpoint");
            renderItems.Add(titleItemOrigin);

            var group = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(group);

            var titleItem = new SvgTitleItem(DrawingRenderer, "DataLabel size adjustment");
            renderItems.Add(titleItem);

            //AppendDebugBounds(renderItems);

            _txtBox.AppendRenderItems(renderItems);
            
            if(_renderConnectionPointLines)
            {
                if (_connectionPointLines != null)
                {
                    _connectionPointLines.AppendRenderItems(renderItems);
                }
            }

            if (_seriesIcon != null)
            {
                var height = Bounds.Height;
                if (height == 0)
                {
                    height = _txtBox.Height;
                }
                //Currently series icon always has a y1 y2 of 2
                var iconGrp = new SvgGroupItem(ChartRenderer, _seriesIcon.Bounds.Left, height / 2 - 2);
                renderItems.Add(iconGrp);
                renderItems.Add(_seriesIcon);
                renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
            }

            if (_leaderLines != null && _leaderLines.Count > 0)
            {
                foreach (var line in _leaderLines)
                {
                    renderItems.Add(line);
                }
            }
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }
    }
}
