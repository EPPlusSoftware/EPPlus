/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.SupportingPropertyBags;
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
            if (r != null)
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
            // this will initialize the metadata, it is needed for the Rich data...

            Structures = new ExcelRichValueStructureCollection(wb, this);
            RichValueRels = new RichValueRelCollection(wb);
            Values = new ExcelRichValueCollection(wb, this);
            SupportingPropertyBagStructures = new SupportingPropertyBagStructureCollection(wb);
            SupportingPropertyBags = new SupportingPropertyBags(wb);
            _richDataDeletions = new ExcelRichDataDeletions();
        }



        internal ExcelRichDataValueTypeInfo ValueTypes { get; }
        internal ExcelRichValueStructureCollection Structures { get; }
        internal ExcelRichValueCollection Values { get; }
        internal RichValueRelCollection RichValueRels { get; }

        internal SupportingPropertyBagStructureCollection SupportingPropertyBagStructures { get; }

        internal SupportingPropertyBags SupportingPropertyBags { get; }

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

        //internal RichValueRel GetRelation(int index)
        //{
        //    return RichValueRels.Items[index];
        //}

        internal RichValueRel GetRelation(string target, string type)
        {
            return RichValueRels.FirstOrDefault(x => x.TargetUri.OriginalString == target && x.Type == type);
        }
    }
}
