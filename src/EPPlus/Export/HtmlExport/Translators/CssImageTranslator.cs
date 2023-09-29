using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssImageTranslator : TranslatorBase
    {
        HtmlImage _p;
        string _encodedImage;
        internal ePictureType? type;

        public CssImageTranslator(HtmlImage p)
        {
            _p = p;
            _encodedImage = ImageEncoder.EncodeImage(p, out type);
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            AddDeclaration("content", $"url('data:{GetContentType(type.Value)};base64,{_encodedImage}')");

            if (context.Pictures.Position != ePicturePosition.DontSet)
            {
               AddDeclaration("position", $"{context.Pictures.Position.ToString().ToLower()}");
            }

            if (_p.FromColumnOff != 0 && context.Pictures.AddMarginLeft)
            {
                var leftOffset = _p.FromColumnOff / ExcelPicture.EMU_PER_PIXEL;
                AddDeclaration("margin-left", $"{leftOffset}px");
            }

            if (_p.FromRowOff != 0 && context.Pictures.AddMarginTop)
            {
                var topOffset = _p.FromRowOff / ExcelPicture.EMU_PER_PIXEL;
                AddDeclaration("margin-top", $"{topOffset}px");
            }

            return declarations;
        }

        private object GetContentType(ePictureType type)
        {
            switch (type)
            {
                case ePictureType.Ico:
                    return "image/vnd.microsoft.icon";
                case ePictureType.Jpg:
                    return "image/jpeg";
                case ePictureType.Svg:
                    return "image/svg+xml";
                case ePictureType.Tif:
                    return "image/tiff";
                default:
                    return $"image/{type}";
            }
        }
    }
}
