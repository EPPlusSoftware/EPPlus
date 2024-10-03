using OfficeOpenXml.Constants;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.Structures;
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
                if (ValueTypes.Part == null)
                {
                    ValueTypes.CreateDefault();
                }
            }
            Structures = new ExcelRichValueStructureCollection(wb);
            Values = new ExcelRichValueCollection(wb, Structures);
            RichValueRels = new RichValueRelCollection(wb);
            _richDataDeletions = new ExcelRichDataDeletions();
        }



        internal ExcelRichDataValueTypeInfo ValueTypes { get; }
        internal ExcelRichValueStructureCollection Structures { get; }
        internal ExcelRichValueCollection Values { get; }
        internal RichValueRelCollection RichValueRels { get; }

        private ExcelRichDataDeletions _richDataDeletions;

        internal ExcelRichDataDeletions Deletions { 
            get 
            {
                return _richDataDeletions;
            } 
        }
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
                RichValueRels.Part.ShouldBeSaved = true;
            }
        }

        internal RichValueRel GetRelation(int index)
        {
            return RichValueRels.Items[index];
        }

        internal RichValueRel GetRelation(string target, string type)
        {
            return RichValueRels.Items.FirstOrDefault(x => x.TargetUri.OriginalString == target && x.Type == type);
        }
    }
}
