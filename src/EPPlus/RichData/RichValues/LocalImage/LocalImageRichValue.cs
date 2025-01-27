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

using OfficeOpenXml.CellPictures;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.LocalImage
{
    internal class LocalImageRichValue : ExcelRichValue
    {

        public LocalImageRichValue(RichDataDatabase richDataDb) : base(richDataDb, RichDataStructureTypes.LocalImage)
        {
            _richDataDb = richDataDb;
        }

        private readonly RichDataDatabase _richDataDb;

        internal override void SetStructure(RichDataDatabase richDataDb)
        {
            base.SetStructure(richDataDb);
            // the first key is a relation and not included in the _keysAndvalues dictionary.
        }

        public Uri ImageUri
        {
            get
            {
                return GetRelation(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier);
            }
            set
            {
                SetRelation(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, "LocalImageIdentifier", value, out uint rvRelId);
                SetValue(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, rvRelId);
                
            }
        }

        public CalcOrigins? CalcOrigin
        {
            get
            {
                var val = GetValueInt(StructureKeyNames.LocalImages.Image.CalcOrigin);
                if(val.HasValue)
                {
                    return (CalcOrigins)val;
                }
                return null;
            }
            set
            {
                SetValue(StructureKeyNames.LocalImages.Image.CalcOrigin, (int?)value);
            }
        }

        public string Text
        {
            get
            {
                return GetValue(StructureKeyNames.LocalImages.Image.Text);
            }
            set
            {
                SetValue(StructureKeyNames.LocalImages.Image.Text, value);
            }
        }

        internal override void PostProcessInitialRead()
        {
            base.PostProcessInitialRead();
            var value = Values.FirstOrDefault(k => k.Key.Name == StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier);
            if(value.ValueInt.HasValue)
            {
                var rvRelId = _richDataDb.RichValueRels.GetIdByIndex(value.ValueInt.Value);
                var rvRel = _richDataDb.RichValueRels.Get(rvRelId);
                SetRelation(value.Key.Name, value.Key.RelationName, rvRel.TargetUri, out uint rid);
                SetValue(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, rvRelId);
            }
            
           
        }
    }
}
