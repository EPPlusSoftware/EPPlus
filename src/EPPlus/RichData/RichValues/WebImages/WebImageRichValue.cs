using OfficeOpenXml.CellPictures;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.WebImages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.WebImages
{
    internal class WebImageRichValue : ExcelRichValue
    {
        public WebImageRichValue(ExcelWorkbook workbook) : this(workbook.IndexStore, workbook.RichData)
        {
            
        }
        public WebImageRichValue(RichDataIndexStore store, ExcelRichData richData) : base(store, richData, RichDataStructureTypes.WebImage)
        {
        }

        public Uri ImageUri
        {
            get
            {
                var img = GetFirstIncomingRelByType<WebImagesSupportingRichData>();
                if(img != null)
                {
                    return img.Blip;
                }
                return null;
            }
            set
            {
                // TODO: should not be imageidentifier
                SetRelation(StructureKeyNames.WebImage.WebImageIdentifier, "LocalImageIdentifier", value);
            }
        }

        /// <summary>
        /// String representation of the alt text associated with the image.
        /// </summary>
        public string Text 
        {
            get 
            {
                return GetValue(StructureKeyNames.WebImage.Text);
            } 
            set 
            {
                SetValue(StructureKeyNames.WebImage.Text, value);
            }
        }

        /// <summary>
        /// Boolean value that when true indicates the image can be generated on demand.
        public bool? ComputedImage
        {
            get
            {
                return GetValueBool(StructureKeyNames.WebImage.ComputedImage);
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ComputedImage, value);
            }
        }

        /// <summary>
        /// Integer value that indicates the dimension modes of the image. There are 4 supported values: 0, 1, 2, 3.
        /// At 0, image fits to the cell and maintains its aspect ratio.
        /// At 1, image fits to the cell and ignore its aspect ratio.
        /// At 2, image maintains its original size and MAY exceed the cell boundary.
        /// At 3, image size is customized by "ImageHeight" and "ImageWidth".
        /// </summary>
        public int? ImageSizing
        {
            get
            {
                var v = GetValueInt(StructureKeyNames.WebImage.ImageSizing);
                return v.HasValue && v >= 0 && v <= 3 ? v : 0;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageSizing, value);
            }
        }

        /// <summary>
        /// Real number representation of the image height in pixels. This property SHOULD only be present when "ImageSizing" is set to 3.
        /// </summary>
        public double? ImageHeight
        {
            get
            {
                if (ImageSizing != 3) return null;
                return GetValueDouble(StructureKeyNames.WebImage.ImageHeight);
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageHeight, value);
            }
        }

        /// <summary>
        /// Real number representation of the image width in pixels. This property SHOULD only be present when "ImageSizing" is set to 3
        /// </summary>
        public double? ImageWidth
        {
            get
            {
                if (ImageSizing != 3) return null;
                return GetValueDouble(StructureKeyNames.WebImage.ImageWidth);
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageWidth, value);
            }
        }

        /// <summary>
        /// Integer value that indicates how the rich value was created.
        /// </summary>
        public CalcOrigins? CalcOrigin
        {
            get
            {
                var val = GetValueInt(StructureKeyNames.WebImage.CalcOrigin);
                if (val.HasValue)
                {
                    return (CalcOrigins)val;
                }
                return null;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.CalcOrigin, (int?)value);
            }
        }
    }
}
