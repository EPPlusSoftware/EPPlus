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

        List<SvgRenderLineItem> LeaderLines = new List<SvgRenderLineItem>();

        Coordinate originPoint = new Coordinate (0, 0);
        Coordinate manualLayoutOffset = new Coordinate (0, 0);

        PointLines lines;
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
        }

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard, SvgTextBox txtBox) : base(chart)
        {
            HasLegendKey = standard.ShowLegendKey;
            TxtBox = txtBox;
            FitToTextBoxContent();
        }

        private void FitToTextBoxContent()
        {
            Bounds.Left = TxtBox.Left;
            Bounds.Top = TxtBox.Top;
            Bounds.Height = TxtBox.Height;
            Bounds.Width = TxtBox.Width;
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

            var separator = string.IsNullOrEmpty(dataLabel.Separator) ? "," : dataLabel.Separator;

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
            ////Reset run y-position.
            ////Datalabel does not use the standard line-spacing textbody offsets
            //txtBox.TextBody.Paragraphs[0].Runs[0].YPosition = 0;

            TxtBox = txtBox;

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
                    ////TopMargin = Rectangle.Top;
                    ////BottomMargin = Rectangle.Bottom;

                    if (dataLabel.ShowLeaderLines)
                    {
                        var cPoints = new ConnectionPointsMiddle(Bounds.Left, Bounds.Top, Bounds.Width, Bounds.Height);

                        //Since this is a child transform changes to this transform will compound
                        lines = new PointLines(ChartRenderer, Bounds, cPoints);
                        //var index = GetClosestConnectionPointToOriginIndex();
                        //ConnectionPointsLines.Clear();

                        ////Add connection points to render
                        //List<string> ptColors = new List<string> { "red", "green", "blue", "yellow" };
                        //for (int i = 0; i < connectionPoints.Points.Count; i++)
                        //{
                        //    var cPoint = connectionPoints.Points[i];
                        //    var cPointLine = new SvgRenderLineItem(chart, txtBox.TextBody.Bounds);
                        //    cPointLine.X1 = 0;
                        //    cPointLine.Y1 = 0;
                        //    cPointLine.X2 = cPoint.X;
                        //    cPointLine.Y2 = cPoint.Y;

                        //    cPointLine.BorderWidth = 1;
                        //    cPointLine.BorderColor = ptColors[i];
                        //    ConnectionPointsLines.Add(cPointLine);
                        //}
                    }

                    //    double xOffset = 0;
                    //    if(index == 0 || index == 2)
                    //    {
                    //        //If Left or Right
                    //        //Add extra 7 px (5.25pt) line to the given side
                    //        var extraLine = new SvgRenderLineItem(chart, Bounds);

                    //        xOffset = index == 0 ? - 5.25d : 5.25d;

                    //        extraLine.X1 = ConnectionPoints[index].X;
                    //        extraLine.Y1 = ConnectionPoints[index].Y;
                    //        extraLine.Y2 = ConnectionPoints[index].Y;
                    //        extraLine.X2 = extraLine.X1 + xOffset;

                    //        extraLine.BorderColor = "black";
                    //        extraLine.BorderWidth = 1;

                    //        LeaderLines.Add(extraLine);
                    //    }
                    //    var mainLine = new SvgRenderLineItem(chart, Bounds);
                    //    mainLine.X1 = ConnectionPoints[index].X + xOffset;
                    //    mainLine.Y1 = ConnectionPoints[index].Y;
                    //    mainLine.X2 = -Bounds.Left;
                    //    mainLine.Y2 = -Bounds.Top;

                    //    mainLine.BorderColor = "black";
                    //    mainLine.BorderWidth = 1;
                    //    LeaderLines.Add(mainLine);
                    //}
                }
            }
            else
            {
                FitToTextBoxContent();
            }
        }

        private int GetClosestConnectionPointToOriginIndex()
        {
            //Origin in local coordinates is 0,0
            return GetClosestConnectionPointCoordinateIndex(new Coordinate(0, 0));
        }

        private int GetClosestConnectionPointCoordinateIndex(Coordinate originPoint)
        {
            //CalculateConnectionPoints();

            double smallestDist = double.MaxValue;
            int i = 0;
            int smallestIndex = 0;

            foreach (var line in lines.RenderLines)
            {
                line.X1 = originPoint.X - line.Bounds.Left - Bounds.Left;
                line.Y1 = originPoint.Y + line.Bounds.Top;

                var w = Math.Abs(line.X2 + manualLayoutOffset.X - originPoint.X);
                var h = Math.Abs(line.Y2 + manualLayoutOffset.Y - originPoint.Y);

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

        internal void SetOriginPointOffset(double xPos, double yPos)
        {
            originPoint.X = xPos;
            originPoint.Y = yPos;

            Bounds.Top = yPos;
            Bounds.Left = xPos;

            if (_hasManualLayout)
            {
                Bounds.Top += manualLayoutOffset.Y;
                Bounds.Left += manualLayoutOffset.X;

                if (_hasLeaderLines)
                {
                    var index = GetClosestConnectionPointCoordinateIndex(originPoint);

                    LeaderLines.Clear();

                    double xOffset = 0;
                    if (index == 0 || index == 2)
                    {
                        //If Left or Right
                        //Add extra 7 px (5.25pt) line to the given side
                        var extraLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);

                        xOffset = index == 0 ? -5.25d : 5.25d;

                        extraLine.X1 = lines.ConnectionPoints.Points[index].X;
                        extraLine.Y1 = lines.ConnectionPoints.Points[index].Y;
                        extraLine.Y2 = lines.ConnectionPoints.Points[index].Y;
                        extraLine.X2 = extraLine.X1 + xOffset;

                        extraLine.BorderColor = "black";
                        extraLine.BorderWidth = 1;

                        LeaderLines.Add(extraLine);
                    }
                    var mainLine = new SvgRenderLineItem(ChartRenderer, ChartRenderer.Bounds);
                    mainLine.X1 = lines.ConnectionPoints.Points[index].X + xOffset;
                    mainLine.Y1 = lines.ConnectionPoints.Points[index].Y;
                    mainLine.X2 = -Bounds.Left;
                    mainLine.Y2 = -Bounds.Top;

                    mainLine.BorderColor = "black";
                    mainLine.BorderWidth = 1;
                    LeaderLines.Add(mainLine);
                }
            }
        }

        //private void CalculateConnectionPoints()
        //{
        //    internalBounds.Left = Bounds.Left;
        //    internalBounds.Top = Bounds.Top;
        //    internalBounds.Width = Bounds.Width;
        //    internalBounds.Height = Bounds.Height;

        //    var middleWidth = Bounds.Width / 2 + LeftMargin;
        //    var middleHeight = Bounds.Height / 2 + TopMargin;

        //    var cPointLeft = new Coordinate(Bounds.Left + LeftMargin , middleHeight);
        //    var cPointTop = new Coordinate(middleWidth, Bounds.Top + TopMargin);
        //    var cPointRight = new Coordinate(Bounds.Right + LeftMargin, middleHeight);
        //    var cPointBottom = new Coordinate(middleWidth, Bounds.Bottom + TopMargin);

        //    ConnectionPoints = new List<Coordinate> { cPointLeft, cPointTop, cPointRight, cPointBottom };
        //}

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var group = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(group);

            var titleItem = new SvgTitleItem(DrawingRenderer, "DataLabelRect");
            renderItems.Add(titleItem);
            SvgRenderRectItem rect = new SvgRenderRectItem(ChartRenderer, Bounds);
            rect.Bounds.Left = 0;
            rect.Bounds.Top = 0;
            rect.Bounds.Width = Bounds.Width;
            rect.Bounds.Height = Bounds.Height;

            rect.FillColor = "red";
            rect.FillOpacity = 0.2;
            renderItems.Add(rect);

            TxtBox.AppendRenderItems(renderItems);

            if (lines != null)
            {
                lines.AppendRenderItems(renderItems);
            }
            //renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));

            //if (LeaderLines != null && LeaderLines.Count > 0)
            //{
            //    foreach (var line in LeaderLines)
            //    {
            //        renderItems.Add(line);
            //    }
            //}
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }
    }
}
