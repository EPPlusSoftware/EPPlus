using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table
{
    internal class RowsDeletedEventArgs : EventArgs
    {
        public RowsDeletedEventArgs(int nRowsDeleted, int position)
        {
            NumberOfDeletedRows = nRowsDeleted;
            Position = position;
        }

        public int NumberOfDeletedRows { get; private set; }

        public int Position { get; set; }
    }
}
