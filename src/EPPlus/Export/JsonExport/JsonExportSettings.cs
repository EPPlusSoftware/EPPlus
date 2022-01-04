using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.JsonExport
{
    public class JsonRangeExportSettings
    {
        public bool FirstRowIsHeader { get; set; } = true;
        public bool AddColumns { get; set; } = true;
    }
}
