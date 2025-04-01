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
using OfficeOpenXml.RichData.RichValueArrays;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.Structures.SupportingPropertyBags;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Types;
using OfficeOpenXml.RichData.WebImages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData.IndexRelations;

namespace OfficeOpenXml.RichData
{
    internal class RichDataDatabase
    {
        public RichDataDatabase(ExcelWorkbook wb, ExcelRichData richData)
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
            IndexStore = wb.IndexStore;
            Structures = new ExcelRichValueStructureCollection(wb, this);
            WebImages = new WebImagesSupportingRichDataCollection(wb);
            RichValueRels = new RichValueRelCollection(wb);
            RichValueValues = new ExcelRichValueValueCollection(wb.IndexStore);
            Values = new ExcelRichValueCollection(wb, this);
            SupportingPropertyBagStructures = new SupportingPropertyBagStructureCollection(wb);
            SupportingPropertyBags = new SupportingPropertyBags(wb);
            RichDataArrayValues = new ExcelRichDataArrayValueCollection(wb.IndexStore);
            RichDataArrays = new ExcelRichDataArrayCollection(wb, this);
        }

        internal RichDataIndexStore IndexStore { get; private set; }

        internal ExcelRichDataValueTypeInfo ValueTypes { get; }
        internal ExcelRichValueStructureCollection Structures { get; }

        internal ExcelRichValueValueCollection RichValueValues { get; set; }
        internal ExcelRichValueCollection Values { get; }
        internal RichValueRelCollection RichValueRels { get; }

        internal SupportingPropertyBagStructureCollection SupportingPropertyBagStructures { get; }

        internal WebImagesSupportingRichDataCollection WebImages { get; set; }

        internal SupportingPropertyBags SupportingPropertyBags { get; }

        internal ExcelRichDataArrayCollection RichDataArrays { get; }

        internal ExcelRichDataArrayValueCollection RichDataArrayValues { get; }
    }
}
