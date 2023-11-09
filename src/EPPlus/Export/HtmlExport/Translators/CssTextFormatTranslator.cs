﻿using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssTextFormatTranslator : CssTableTextFormatTranslator
    {
        bool _wrapText;
        int _indent;
        int _textRotation;

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

            if (_applyAlignment ?? false)
            {
                var hAlign = GetHorizontalAlignment();
                var vAlign = GetVerticalAlignment();

                if (string.IsNullOrEmpty(hAlign) == false && context.Exclude.HorizontalAlignment == false)
                {
                    AddDeclaration("text-align", hAlign);
                }

                if (_verticalAlignment != ExcelVerticalAlignment.Bottom && context.Exclude.VerticalAlignment == false)
                {
                    AddDeclaration("vertical-align", vAlign);
                }
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

        //protected string GetVerticalAlignment()
        //{
        //    switch (_verticalAlignment)
        //    {
        //        case ExcelVerticalAlignment.Top:
        //            return "top";
        //        case ExcelVerticalAlignment.Center:
        //            return "middle";
        //        case ExcelVerticalAlignment.Bottom:
        //            return "bottom";
        //    }

        //    return "";
        //}

        //protected string GetHorizontalAlignment()
        //{
        //    switch (_horizontalAlignment)
        //    {
        //        case ExcelHorizontalAlignment.Right:
        //            return "right";
        //        case ExcelHorizontalAlignment.Center:
        //        case ExcelHorizontalAlignment.CenterContinuous:
        //            return "center";
        //        case ExcelHorizontalAlignment.Left:
        //            return "left";
        //    }

        //    return "";
        //}


    }
}
