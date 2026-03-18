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
    internal class SvgChartDataLabelStandard : SvgChartObject
    {
        internal bool HasLegendKey { get; private set; } = false;

        bool _hasManualLayout = false;
        bool _hasLeaderLines = false;
        bool haveAdjustedForIcon = false;

        internal SvgTextBox TxtBox;

        BoundingBox _parentPoint;

        List<SvgRenderLineItem> LeaderLines = new List<SvgRenderLineItem>();

        Coordinate _manualLayoutOffset = new Coordinate (0, 0);
        bool renderConnectionPointLines = false;

        PointLines _connectionPointLines;
        eLabelPosition _labelPosition;

        public SvgChartDataLabelStandard(DrawingChart chart, string dataLabelText) : base(chart)
        {
            var txtBox = new SvgTextBox(chart, chart.Bounds, chart.Bounds);
            txtBox.AddText(0, dataLabelText);
        }

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard) : base(chart)
        {
            HasLegendKey = standard.ShowLegendKey;
            _labelPosition = standard.Position;
        }

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard, SvgTextBox txtBox) : base(chart)
        {
            HasLegendKey = standard.ShowLegendKey;
            TxtBox = txtBox;
        }

        RenderItem _seriesIcon = null;

        internal void AddSeriesIcon(RenderItem seriesIcon)
        {
            var iconWidth = seriesIcon.Bounds.Width;
            var iconHeight = seriesIcon.Bounds.Height;

            _seriesIcon = seriesIcon;
            _seriesIcon.Bounds.Parent = _parentPoint;

            if (haveAdjustedForIcon == false)
            {
                if (_hasManualLayout == false)
                {
                    TxtBox.Left += iconWidth + TxtBox.LeftMargin;
                    if (iconHeight > TxtBox.Height)
                    {
                        Bounds.Height = iconHeight;
                    }
                }
                else
                {
                    Bounds.Left += iconWidth + TxtBox.LeftMargin;
                    Bounds.Width += iconWidth;
                    Bounds.Height += iconHeight;
                }
                //It seems there is a hard-coded margin in excel of about 4.5pt (6px) in addition to the width of the marker
                Bounds.Left += 4.5d;
                haveAdjustedForIcon = true;
            }
        }

        internal void ImportDataLabel(SvgChart chart, ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, ExcelDrawingParagraph defaultParagraph, BoundingBox maxBounds)
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
                dlblStrings.Add(yValue.ToString());
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

            //Center the textbox at the origin point
            Bounds.Left -= txtBox.Width / 2;
            Bounds.Top -= txtBox.Height / 2;

            TxtBox = txtBox;

            if (dataLabel.Fill.IsEmpty == false)
            {
                TxtBox.Rectangle.FillColor = "#" + dataLabel.Fill.Color.ToColorString();
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
                    TxtBox.Rectangle.FillColor = "#" + individualLabel.Fill.Color.ToColorString();
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

                    if (dataLabel.ShowLeaderLines)
                    {
                        var cPoints = new ConnectionPointsMiddle(TxtBox.Rectangle.Bounds.Left, TxtBox.Rectangle.Bounds.Top, TxtBox.Rectangle.Bounds.Width, TxtBox.Rectangle.Bounds.Height);

                        //Since this is a child transform changes to this transform will compound
                        _connectionPointLines = new PointLines(ChartRenderer, Bounds, cPoints);
                    }
                }
            }
        }

        private int GetClosestConnectionPointCoordinateIndex(Coordinate originPoint)
        {
            //CalculateConnectionPoints();

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

        internal void SetParentPoint(BoundingBox parentPoint)
        {
            Bounds.Parent = parentPoint;
            _parentPoint = parentPoint;

            switch (_labelPosition)
            {
                case eLabelPosition.Center:
                    break;
                case eLabelPosition.Left:
                    Bounds.Left -= TxtBox.Width + (parentPoint.Width / 2);
                    break;
                case eLabelPosition.Right:
                case eLabelPosition.BestFit:
                    Bounds.Left += TxtBox.Width / 2 + parentPoint.Width;
                    break;
                case eLabelPosition.Top:
                    Bounds.Top -= (parentPoint.Height + TxtBox.Height) / 2;
                    break;
                case eLabelPosition.Bottom:
                    Bounds.Top += (parentPoint.Height + TxtBox.Height) / 2;
                    break;
                default:
                    throw new InvalidOperationException($"The datalabel position {_labelPosition} has not been implemented yet");
            }

            if (_hasManualLayout)
            {
                if (_hasLeaderLines)
                {
                    _connectionPointLines.UpdateLines();

                    var originPoint = new Coordinate(-Bounds.Left, -Bounds.Top);

                    var index = GetClosestConnectionPointCoordinateIndex(originPoint);

                    LeaderLines.Clear();

                    double xOffset = 0;
                    if (index == 0 || index == 2)
                    {
                        //If Left or Right
                        //Add extra 7 px (5.25pt) line to the given side
                        var extraLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);

                        xOffset = index == 0 ? -5.25d : 5.25d;

                        extraLine.X1 = _connectionPointLines.ConnectionPoints.Points[index].X;
                        extraLine.Y1 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.Y2 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                        extraLine.X2 = extraLine.X1 + xOffset;

                        extraLine.BorderColor = "gray";
                        extraLine.BorderWidth = 0.5;

                        LeaderLines.Add(extraLine);
                    }
                    var mainLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);
                    mainLine.X1 = _connectionPointLines.ConnectionPoints.Points[index].X + xOffset;
                    mainLine.Y1 = _connectionPointLines.ConnectionPoints.Points[index].Y;
                    mainLine.X2 = originPoint.X;
                    mainLine.Y2 = originPoint.Y;

                    mainLine.BorderColor = "gray";
                    mainLine.BorderWidth = 0.5;
                    LeaderLines.Add(mainLine);
                }
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var parentPointGroup = new SvgGroupItem(ChartRenderer, _parentPoint);
            renderItems.Add(parentPointGroup);

            var titleItemOrigin = new SvgTitleItem(DrawingRenderer, "DataLabel originpoint");
            renderItems.Add(titleItemOrigin);

            //SvgRenderRectItem markerRect = new SvgRenderRectItem(DrawingRenderer, _parentPoint);

            //markerRect.Width = _parentPoint.Width;
            //markerRect.Height = _parentPoint.Height;

            //markerRect.Top -= _parentPoint.Height / 2;
            //markerRect.Left -= _parentPoint.Width / 2;

            //markerRect.FillColor = "blue";
            //markerRect.FillOpacity = 0.2d;

            //renderItems.Add(markerRect);

            var group = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(group);

            var titleItem = new SvgTitleItem(DrawingRenderer, "DataLabel size adjustment");
            //renderItems.Add(titleItem);
            //SvgRenderRectItem rect = new SvgRenderRectItem(ChartRenderer, Bounds);
            //rect.Bounds.Left = 0;
            //rect.Bounds.Top = 0;
            //rect.Bounds.Width = Bounds.Width;
            //rect.Bounds.Height = Bounds.Height;

            //rect.FillColor = "red";
            //rect.FillOpacity = 0.2;
            //renderItems.Add(rect);

            //TxtBox.Rectangle.FillColor = "red";
            //TxtBox.Rectangle.FillOpacity = 1d;

            if(_seriesIcon != null)
            {
                var height = Bounds.Height;
                if(height == 0)
                {
                    height = TxtBox.Height;
                }
                //Currently series icon always has a y1 y2 of 2
                var iconGrp = new SvgGroupItem(ChartRenderer, 0, height / 2 - 2);
                renderItems.Add(iconGrp);
                renderItems.Add(_seriesIcon);
                renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
            }

            TxtBox.AppendRenderItems(renderItems);
            
            if(renderConnectionPointLines)
            {
                if (_connectionPointLines != null)
                {
                    _connectionPointLines.AppendRenderItems(renderItems);
                }
            }

            if (LeaderLines != null && LeaderLines.Count > 0)
            {
                foreach (var line in LeaderLines)
                {
                    renderItems.Add(line);
                }
            }
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }
    }
}
