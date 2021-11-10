using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class TableStyleToCss
    {
        ExcelTable _table;
        internal TableStyleToCss(ExcelTable table)
        {
            _table = table;
        }
        internal void Render(StreamWriter sw)
        {
            if(_table.TableStyle==TableStyles.None)
            {
                return;
            }
        }


    }
}
