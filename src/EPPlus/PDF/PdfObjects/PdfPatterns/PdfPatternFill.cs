using OfficeOpenXml.PDF.PdfGraphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal abstract class PdfPatternFill
    {
        public PdfColor Background;
        public PdfColor Foreground;
        private readonly List<string> commands = new List<string>();

        protected PdfPatternFill(PdfColor foreground, PdfColor background)
        {
            Foreground = foreground;
            Background = background;
        }

        public abstract string CreatePatternResource();
    }
}
