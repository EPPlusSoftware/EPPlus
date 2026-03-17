using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

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

        Coordinate originPoint = new Coordinate (0, 0);
        Coordinate manualLayoutOffset = new Coordinate (0, 0);

        PointLines connectionPointLines;

        eLabelPosition _labelPosition;
        ////Connection point coords are accurate to Internal bounds
        //BoundingBox internalBounds = new BoundingBox();
        //List<Coordinate> ConnectionPoints = new List<Coordinate>();

        public SvgChartDataLabelStandard(DrawingChart chart, string dataLabelText) : base(chart)
        {
            var txtBox = new SvgTextBox(chart, chart.Bounds, chart.Bounds);
            txtBox.AddText(0, dataLabelText);
            FitToTextBoxContent();
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
            FitToTextBoxContent();
        }

        private void FitToTextBoxContent()
        {
            //Bounds.Left = TxtBox.Left;
            //Bounds.Top = TxtBox.Top;
            //Bounds.Height = TxtBox.Height;
            //Bounds.Width = TxtBox.Width;
        }

        internal void AddSeriesIcon(double iconWidth, double iconHeight)
        {
            if (haveAdjustedForIcon == false)
            {
                if (_hasManualLayout == false)
                {
                    TxtBox.Left += iconWidth + TxtBox.LeftMargin;
                    FitToTextBoxContent();
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

            //txtBox.LeftMargin = 3.1181d;
            //txtBox.RightMargin = 3.1181d;
            //txtBox.TopMargin = 1.4173d;
            //txtBox.BottomMargin = 1.4173d;

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

            Bounds.Left -= txtBox.Width / 2;
            Bounds.Top -= txtBox.Height / 2;
            ////Reset run y-position.
            ////Datalabel does not use the standard line-spacing textbody offsets
            //txtBox.TextBody.Paragraphs[0].Runs[0].YPosition = 0;

            TxtBox = txtBox;

            _labelPosition = dataLabel.Position;

            //TxtBox.TextBody.Bounds.Left += TxtBox.LeftMargin;
            //TxtBox.TextBody.Bounds.Top += TxtBox.TopMargin;

            //TxtBox.Left = txtBox.LeftMargin;
            //TxtBox.Top = txtBox.TopMargin;

            if (dataLabel is ExcelChartDataLabelItem)
            {
                var individualLabel = dataLabel as ExcelChartDataLabelItem;

                if (individualLabel.Layout != null && individualLabel.Layout.HasLayout)
                {
                    FitToTextBoxContent();
                    _hasManualLayout = true;
                    var rect = GetRectFromManualLayout(chart, individualLabel.Layout);
                    Rectangle = rect;

                    //LeftMargin = Rectangle.Left;
                    //TopMargin = Rectangle.Top;
                    manualLayoutOffset = new Coordinate(Rectangle.Left, Rectangle.Top);

                    if (dataLabel.ShowLeaderLines)
                    {
                        _hasLeaderLines = true;
                    }

                    if (dataLabel.ShowLeaderLines)
                    {
                        var cPoints = new ConnectionPointsMiddle(Bounds.Left, Bounds.Top, Bounds.Width, Bounds.Height);

                        //Since this is a child transform changes to this transform will compound
                        connectionPointLines = new PointLines(ChartRenderer, Bounds, cPoints);
                    }
                }
            }
            else
            {
                FitToTextBoxContent();
            }
        }

        private int GetClosestConnectionPointCoordinateIndex(Coordinate originPoint)
        {
            //CalculateConnectionPoints();

            double smallestDist = double.MaxValue;
            int i = 0;
            int smallestIndex = 0;

            foreach (var line in connectionPointLines.RenderLines)
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
        }

        internal void SetOriginPointOffset(double xPos, double yPos)
        {
            originPoint.X = xPos;
            originPoint.Y = yPos;

            Bounds.Top = yPos;
            Bounds.Left = xPos;

            if (_hasManualLayout)
            {
                //Bounds.Top += manualLayoutOffset.Y;
                //Bounds.Left += manualLayoutOffset.X;

                //if (_hasLeaderLines)
                //{
                //    originPoint.Y -= Bounds.Top;
                //    originPoint.X -= Bounds.Left;

                //    var index = GetClosestConnectionPointCoordinateIndex(originPoint);

                //    LeaderLines.Clear();

                //    double xOffset = 0;
                //    if (index == 0 || index == 2)
                //    {
                //        //If Left or Right
                //        //Add extra 7 px (5.25pt) line to the given side
                //        var extraLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);

                //        xOffset = index == 0 ? -5.25d : 5.25d;

                //        extraLine.X1 = connectionPointLines.ConnectionPoints.Points[index].X;
                //        extraLine.Y1 = connectionPointLines.ConnectionPoints.Points[index].Y;
                //        extraLine.Y2 = connectionPointLines.ConnectionPoints.Points[index].Y;
                //        extraLine.X2 = extraLine.X1 + xOffset;

                //        extraLine.BorderColor = "gray";
                //        extraLine.BorderWidth = 0.5;

                //        LeaderLines.Add(extraLine);
                //    }
                //    var mainLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);
                //    mainLine.X1 = connectionPointLines.ConnectionPoints.Points[index].X + xOffset;
                //    mainLine.Y1 = connectionPointLines.ConnectionPoints.Points[index].Y;
                //    mainLine.X2 = originPoint.X;
                //    mainLine.Y2 = originPoint.Y;

                //    mainLine.BorderColor = "gray";
                //    mainLine.BorderWidth = 0.5;
                //    LeaderLines.Add(mainLine);
                //}
            }
        }

        bool renderConnectionPointLines = false;

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var parentPointGroup = new SvgGroupItem(ChartRenderer, _parentPoint);
            renderItems.Add(parentPointGroup);

            var titleItemOrigin = new SvgTitleItem(DrawingRenderer, "DataLabel originpoint");
            renderItems.Add(titleItemOrigin);

            SvgRenderRectItem markerRect = new SvgRenderRectItem(DrawingRenderer, _parentPoint);

            markerRect.Width = _parentPoint.Width;
            markerRect.Height = _parentPoint.Height;

            markerRect.Top -= _parentPoint.Height / 2;
            markerRect.Left -= _parentPoint.Width / 2;

            markerRect.FillColor = "blue";
            markerRect.FillOpacity = 0.2d;

            renderItems.Add(markerRect);

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

            TxtBox.Rectangle.FillColor = "red";
            TxtBox.Rectangle.FillOpacity = 1d;

            TxtBox.AppendRenderItems(renderItems);
            
            if(renderConnectionPointLines)
            {
                if (connectionPointLines != null)
                {
                    connectionPointLines.AppendRenderItems(renderItems);
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
