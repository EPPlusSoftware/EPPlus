using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    public class LoadFromDataTableParams
    {
        public bool PrintHeaders { get; set; }

        public TableStyles TableStyle { get; set; }
    }
}
