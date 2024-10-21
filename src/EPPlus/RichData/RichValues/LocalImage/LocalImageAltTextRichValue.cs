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
    internal class LocalImageAltTextRichValue : ExcelRichValue
    {
        public LocalImageAltTextRichValue(ExcelWorkbook workbook) : this(workbook.IndexStore, workbook.RichData)
        {
        }

        public LocalImageAltTextRichValue(RichDataIndexStore store, ExcelRichData richData) : base(store, richData, RichDataStructureTypes.LocalImageWithAltText)
        {
        }

        public Uri ImageUri
        {
            get
            {
                return GetRelation(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier);
            }
            set
            {
                SetRelation(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier, "LocalImageIdentifier", value);
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
