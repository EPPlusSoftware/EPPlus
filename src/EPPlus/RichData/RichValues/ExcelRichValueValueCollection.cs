using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelRichValueValueCollection : IndexedCollection<ExcelRichValueValue>
    {
        public ExcelRichValueValueCollection(RichDataIndexStore store) : base(store, RichDataEntities.RichValueValue)
        {
        }
    }
}
