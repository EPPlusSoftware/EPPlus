using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions
{
    internal enum RelativeTo
    {
        ColumnTotals = 0,
        RowTotals = 1,
        GrandTotals = 2,
        ParentColTotal = 3,
        ParentRowTotal = 4
    }
}
