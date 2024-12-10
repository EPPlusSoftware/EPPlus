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
using OfficeOpenXml.RichData.RichValueArrays;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.SupportingPropertyBags;
using OfficeOpenXml.RichData.Types;
using OfficeOpenXml.RichData.WebImages;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichData
    {
        internal ExcelRichData(ExcelWorkbook wb)
        {
            Db = new RichDataDatabase(wb, this);
            _richDataDeletions = new ExcelRichDataDeletions();
        }



        internal RichDataDatabase Db { get; private set; }

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
            Db.ValueTypes.CreatePart();
            Db.Structures.CreatePart();
            Db.Values.CreatePart();
            Db.RichDataArrays.CreatePart();
        }

        internal void SetHasValuesOnParts()
        {
            if(Db.ValueTypes.Part.ShouldBeSaved==false)
            {
                Db.ValueTypes.Part.ShouldBeSaved = true;
                Db.Structures.Part.ShouldBeSaved = true;
                Db.Values.Part.ShouldBeSaved = true;
                Db.RichValueRels.Part.ShouldBeSaved = true;
                Db.RichDataArrays.Part.ShouldBeSaved = true;
            }
        }

        internal RichValueRel GetRelation(string target, string type)
        {
            return Db.RichValueRels.FirstOrDefault(x => x.TargetUri.OriginalString == target && x.Type == type);
        }
    }
}
