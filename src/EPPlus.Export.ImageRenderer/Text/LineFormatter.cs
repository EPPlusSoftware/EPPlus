using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal static class LineFormatter
    {
        internal static List<string> GetFormattedLines(ExcelTextBody body, out List<double> lineWidths)
        {
            var paragraphs = body.Paragraphs;
            var text = paragraphs.Text;

            List<int> txtRunIndicies = new List<int>();
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


        }

        //internal static List<string> WrapLines(ExcelDrawingTextRunCollection txtRuns, double maxWidth)
        //{
        //    txtRun
        //    foreach (var run in txtRuns)
        //    {

        //    }
        //}
    }
}


