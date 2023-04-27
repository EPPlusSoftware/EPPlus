using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData.Types;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichData
    {
        internal ExcelRichData(ExcelWorkbook wb)
        {
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueTypeRelationship).FirstOrDefault();
            if(r != null)
            {
                ValueTypes = new ExcelRichDataValueTypeInfo(wb, r);
                Structures = new ExcelRichValueStructureCollection(wb);
                Values = new ExcelRichValueCollection(wb, Structures);
            }
            else
            {
                ValueTypes.CreateDefault();
            }
        }
        internal ExcelRichDataValueTypeInfo ValueTypes { get; }
        internal ExcelRichValueStructureCollection Structures { get; }
        internal ExcelRichValueCollection Values { get; }
        internal void Save()
        {
            ValueTypes.Save();
            Structures.Save();
            Values.Save();
        }
    }
}
