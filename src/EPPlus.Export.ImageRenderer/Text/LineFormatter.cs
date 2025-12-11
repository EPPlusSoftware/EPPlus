using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal static class LineFormatter
    {
        internal static List<int> GetTextRunIndicies(ExcelTextBody body)
        {
            var paragraphs = body.Paragraphs;
            var text = paragraphs.Text;
            List<int> txtRunIndicies = new List<int>();
            List<double> AdvanceWidths = new List<double> ();

            int lastIndex = 0;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var p = paragraphs[i];

                for (int j = 0; j < p.TextRuns.Count; j++)
                {
                    var run = p.TextRuns[j];

                    var indexOfRun = text.IndexOf(run.Text, lastIndex);
                    txtRunIndicies.Add(indexOfRun);
                    lastIndex = indexOfRun;
                }
            }



            return txtRunIndicies;
        }


        //internal static List<string> GetFormattedLines(ExcelTextBody body, out List<double> lineWidths, double MaxWidth = double.NaN)
        //{
        //    var paragraphs = body.Paragraphs;
        //    var text = paragraphs.Text;

        //    List<int> txtRunIndicies = new List<int>();
        //    List<MeasurementFont> fonts = new List<MeasurementFont>();

        //    int lastIndex = 0;

        //    FontMeasurerTrueType measurer = new FontMeasurerTrueType();

        //    for (int i = 0; i < paragraphs.Count; i++)
        //    {
        //        var p = paragraphs[i];

        //        for (int j = 0; j < p.TextRuns.Count; j++)
        //        {
        //            var run = p.TextRuns[j];
        //            TextWrapper.GetLines()

        //            fonts.Add(run.GetMeasurementFont());
        //            var indexOfRun = text.IndexOf(run.Text, lastIndex);
        //            txtRunIndicies.Add(indexOfRun);
        //            lastIndex = indexOfRun;
        //        }
        //    }



        //    //return List<string> test = new List<string>();
        //    //foreach (var font in fonts)
        //    //{
        //    //    FontMeasurerTrueType measurer = new FontMeasurerTrueType();
        //    //    measurer.SetFont(font);


        //    //}

        //    //}

        //    //internal static List<string> WrapLines(ExcelDrawingTextRunCollection txtRuns, double maxWidth)
        //    //{
        //    //    txtRun
        //    //    foreach (var run in txtRuns)
        //    //    {

        //    //    }
        //    //}
        //}
    }
}


