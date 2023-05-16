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
            }
            else
            {
                ValueTypes = new ExcelRichDataValueTypeInfo(wb);
                ValueTypes.CreateDefault();
            }
            Structures = new ExcelRichValueStructureCollection(wb);
            Values = new ExcelRichValueCollection(wb, Structures);
        }
        internal ExcelRichDataValueTypeInfo ValueTypes { get; }
        internal ExcelRichValueStructureCollection Structures { get; }
        internal ExcelRichValueCollection Values { get; }
        internal void CreateParts()
        {
            //Creates the rich data parts and add the parts to the package. 
            //As richtext depends on the worksheet to be saved to get value and cell meta data depending on rich data, it is save using a save handler.
            ValueTypes.CreatePart();
            Structures.CreatePart();
            Values.CreatePart();
        }

        internal void SetHasValuesOnParts()
        {
            if(ValueTypes.Part.ShouldBeSaved==false)
            {
                ValueTypes.Part.ShouldBeSaved = true;
                Structures.Part.ShouldBeSaved = true;
                Values.Part.ShouldBeSaved = true;
            }
        }
    }
}
