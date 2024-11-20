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
    //internal class LocalImageAltTextRichValue : ExcelRichValue
    //{
    //    public LocalImageAltTextRichValue(ExcelWorkbook workbook) : this(workbook.IndexStore, workbook.RichData)
    //    {
    //    }

    //    public LocalImageAltTextRichValue(RichDataIndexStore store, ExcelRichData richData) : base(store, richData, RichDataStructureTypes.LocalImageWithAltText)
    //    {
    //    }

    //    public Uri ImageUri
    //    {
    //        get
    //        {
    //            return GetRelation(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier);
    //        }
    //        set
    //        {
    //            SetRelation(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier, "LocalImageIdentifier", value);
    //        }
    //    }

    //    public CalcOrigins? CalcOrigin
    //    {
    //        get
    //        {
    //            var val = GetValueInt(StructureKeyNames.LocalImages.ImageAltText.CalcOrigin);
    //            if (val.HasValue)
    //            {
    //                return (CalcOrigins)val;
    //            }
    //            return null;
    //        }
    //        set
    //        {
    //            if(value.HasValue)
    //            {
    //                SetValue(StructureKeyNames.LocalImages.ImageAltText.CalcOrigin, (int?)value);
    //            }
                
    //        }
    //    }

    //    public string Text
    //    {
    //        get
    //        {
    //            return GetValue(StructureKeyNames.LocalImages.ImageAltText.Text);
    //        }
    //        set
    //        {
    //            SetValue(StructureKeyNames.LocalImages.ImageAltText.Text, value);
    //        }
    //    }
    //}
}
