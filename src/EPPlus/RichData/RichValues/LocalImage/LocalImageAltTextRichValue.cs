using OfficeOpenXml.CellPictures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.LocalImage
{
    internal class LocalImageAltTextRichValue : ExcelRichValue
    {
        public LocalImageAltTextRichValue(ExcelWorkbook workbook) : base(workbook, RichDataStructureTypes.LocalImageWithAltText)
        {
        }

        public int? RelLocalImageIdentifier
        {
            get
            {
                return GetValueInt(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier);
            }
            set
            {
                SetValue(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier, value);
            }
        }

        public CalcOrigins? CalcOrigin
        {
            get
            {
                var val = GetValueInt(StructureKeyNames.LocalImages.ImageAltText.CalcOrigin);
                if (val.HasValue)
                {
                    return (CalcOrigins)val;
                }
                return null;
            }
            set
            {
                if(value.HasValue)
                {
                    SetValue(StructureKeyNames.LocalImages.ImageAltText.CalcOrigin, (int?)value);
                }
                
            }
        }

        public string Text
        {
            get
            {
                return GetValue(StructureKeyNames.LocalImages.ImageAltText.Text);
            }
            set
            {
                SetValue(StructureKeyNames.LocalImages.ImageAltText.Text, value);
            }
        }
    }
}
