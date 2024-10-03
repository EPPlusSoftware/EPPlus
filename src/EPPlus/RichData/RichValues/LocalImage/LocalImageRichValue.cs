﻿/*************************************************************************************************
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
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.LocalImage
{
    internal class LocalImageRichValue : ExcelRichValue
    {
        public LocalImageRichValue(ExcelWorkbook workbook) : base(workbook, RichDataStructureTypes.LocalImage)
        {
        }

        public int? RelLocalImageIdentifier
        {
            get
            {
                return GetValueInt(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier);
            }
            set
            {
                SetValue(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, value);
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
    }
}