/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System.Collections.Generic;

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
