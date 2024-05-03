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
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssTextFormatTranslator : CssTableTextFormatTranslator
    {
        bool _wrapText;
        int _indent;
        int _textRotation;
        bool _rightDefault;

        internal CssTextFormatTranslator(StyleXml xfs) : base(xfs)
        {
            _wrapText = xfs._style.WrapText;
            _indent = xfs._style.Indent;
            _textRotation = xfs._style.TextRotation;

            _applyAlignment = xfs._style.ApplyAlignment;

            _horizontalAlignment = xfs._style.HorizontalAlignment;
            _verticalAlignment = xfs._style.VerticalAlignment;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            if (context.Exclude.WrapText == false)
            {
                AddDeclaration("white-space", _wrapText ? " break-spaces" : " nowrap");
            }

            var hAlign = "";
            var vAlign = "";

            if ((_horizontalAlignment != ExcelHorizontalAlignment.General) && (context.Exclude.HorizontalAlignment == false))
            {
                hAlign = GetHorizontalAlignment();
            }

            if ((_verticalAlignment != ExcelVerticalAlignment.Center) && (context.Exclude.VerticalAlignment == false))
            {
                vAlign = GetVerticalAlignment();
            }

            if ((string.IsNullOrEmpty(hAlign)) && _rightDefault)
            {
                hAlign = "right";
            }

            if ((string.IsNullOrEmpty(hAlign) == false) && (context.Exclude.HorizontalAlignment == false))
            {
                AddDeclaration("text-align", hAlign);
            }

            if (_verticalAlignment != ExcelVerticalAlignment.Center && (context.Exclude.VerticalAlignment == false))
            {
                AddDeclaration("vertical-align", vAlign);
            }

            if (_textRotation != 0 && context.Exclude.TextRotation == false)
            {
                if (_textRotation == 255)
                {
                    AddDeclaration("writing-mode", "vertical-lr");
                    AddDeclaration("text-orientation", "upright");
                }
                else
                {
                    var rotationvalue = _textRotation > 90 ? _textRotation - 90 : 360 - _textRotation;
                    AddDeclaration("transform", $"rotate({rotationvalue}deg)");
                }
            }
            if (_indent > 0 && context.Exclude.Indent == false)
            {
                AddDeclaration("padding-left", $"{_indent * context.IndentValue}{context.IndentUnit}");
            }

            return declarations;
        }
    }
}
