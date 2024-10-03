using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures
{
    internal static class StructureExtensions
    {
        public static string[] ToNameArray(this List<ExcelRichValueStructureKey> list)
        {
            if (list == null) return null;
            return list.Select(x => x.Name).ToArray();
        }
    }
}
