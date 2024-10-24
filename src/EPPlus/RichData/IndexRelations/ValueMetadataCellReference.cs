using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class ValueMetadataCellReference
    {
        public ValueMetadataCellReference(int worksheetIx, int row, int col)
        {
            WorksheetIx = worksheetIx;
            Row = row;
            Column = col;
        }

        public int WorksheetIx { get; }

        public int Row { get; }

        public int Column { get; }
    }
}
