using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class HtmlDrawing
    {
        public int WorksheetId { get; set; }
        public int FromRow { get; set; }
        public int FromRowOff { get; set; }
        public int ToRow { get; set; }
        public int ToRowOff { get; set; }
        public int FromColumn { get; set; }
        public int FromColumnOff { get; set; }
        public int ToColumn { get; set; }
        public int ToColumnOff { get; set; }
    }
}
