using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexRelation
    {
        public IndexPointer From { get; set; }

        public IndexedValue To { get; set; }

        public int NumberOfPointers { get; set; }
    }
}
