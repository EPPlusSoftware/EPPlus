using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal static class LineFormatter
    {
        internal static /*List<TextLine>*/ void GetFormattedLines(ExcelTextBody body)
        {
            var paragraphs = body.Paragraphs;
            //
            var text = paragraphs.Text;
            for (int i = 0; i < paragraphs.Count; i++) 
            {
                var p = paragraphs[i];
                for(int j = 0; j < p.TextRuns.Count; j++)
                {
                    var run = p.TextRuns[j];
                    run.SplitIntoLines();
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
