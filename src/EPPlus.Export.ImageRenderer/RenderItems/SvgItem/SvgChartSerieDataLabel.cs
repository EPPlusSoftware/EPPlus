using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartSerieDataLabel : SvgChartObject
    {
        internal List<SvgTextBox> dlblTextBoxes = new List<SvgTextBox>();

        string separator;

        public SvgChartSerieDataLabel(SvgChart chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues) : base(chart)
        {
            if (dlblSerie.DataLabels.Count == 0 && serie.NumberOfItems > 0)
            {
                separator = string.IsNullOrEmpty(dlblSerie.Separator) ? "," : dlblSerie.Separator;

                //string dlblStr = "";

                ////dlblStr +=""
                ////StringBuilder sb = new StringBuilder();
                //if (dlblSerie.ShowSeriesName)
                //{
                //    dlblStr += serie.Header;
                //    //AppendToDatalabelStr(sb, serie.Header);
                //}

                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    List<string> dlblStrings = new List<string>();

                    if (dlblSerie.ShowSeriesName)
                    {
                        dlblStrings.Add(serie.GetHeaderString());
                    }
                    if (dlblSerie.ShowCategory)
                    {
                        dlblStrings.Add(xValues[i].ToString());
                    }
                    if (dlblSerie.ShowValue)
                    {
                        dlblStrings.Add(yValues[i].ToString());
                    }

                    string finalString = "";
                    for (int j = 0; j < dlblStrings.Count; j++)
                    {
                        finalString += dlblStrings[j];
                        if (j != dlblStrings.Count - 1)
                        {
                            finalString += separator;
                        }
                    }

                    var txtBox = new SvgTextBox(chart, maxBounds, maxBounds);
                    txtBox.ImportTextBody(dlblSerie.TextBody);

                    txtBox.TextBody.ImportParagraph(dlblSerie.TextBody.Paragraphs[0], 0, finalString);
                    //Remove dummy paragraph added by ImportTextBody
                    txtBox.TextBody.Paragraphs.RemoveAt(0);
                    //Reset run y-position.
                    //Datalabel does not use the standard line-spacing textbody offsets
                    txtBox.TextBody.Paragraphs[0].Runs[0].YPosition = 0;

                    dlblTextBoxes.Add(txtBox);
                }
            }
            else
            {
                foreach (var dlbl in dlblSerie.DataLabels)
                {
                    var txtBox = new SvgTextBox(chart, maxBounds, maxBounds);
                    txtBox.ImportTextBody(dlbl.TextBody);
                    dlblTextBoxes.Add(txtBox);
                }
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            SvgGroupItem groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            renderItems.Add(groupItem);

            foreach (var renderItem in dlblTextBoxes)
            {
                renderItem.AppendRenderItems(renderItems);
            }

            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }

        //internal void AppendToDatalabelStr(StringBuilder sb, string toAppend)
        //{
        //    sb.Append(toAppend);
        //    sb.Append(separator);
        //}

        //internal string GetDatalabelString(ExcelChartSerieDataLabel dlbl, ExcelChartStandardSerie serie)
        //{
        //    var dlblStr = "";

        //    //if(dlbl.ShowValue)
        //    //{
               
        //    //}
        //    //if(dlbl.ShowSeriesName)
        //    //{
        //    //    dlb
        //    //}
        //}

        //public override RenderItemType Type => throw new NotImplementedException();

        //public override void Render(StringBuilder sb)
        //{
        //    throw new NotImplementedException();
        //}
    }
}
