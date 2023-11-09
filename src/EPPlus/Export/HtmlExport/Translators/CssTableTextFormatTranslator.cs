using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssTableTextFormatTranslator : TranslatorBase
    {
        protected ExcelHorizontalAlignment _horizontalAlignment;
        protected ExcelVerticalAlignment _verticalAlignment;
        protected bool? _applyAlignment;
        bool _rightDefault = false;

        internal CssTableTextFormatTranslator(StyleXml xfs, bool rightDefault = false)
        {
            _applyAlignment = xfs._style.ApplyAlignment;

            _horizontalAlignment = xfs._style.HorizontalAlignment;
            _verticalAlignment = xfs._style.VerticalAlignment;
            _rightDefault = rightDefault;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            var hAlign = "";
            var vAlign = "";

            if (_applyAlignment ?? false)
            {
                hAlign = GetHorizontalAlignment();
                vAlign = GetVerticalAlignment();
            }

            if(string.IsNullOrEmpty(hAlign) && _rightDefault)
            {
                hAlign = "right";
            }

            if (!string.IsNullOrEmpty(hAlign) && context.Exclude.HorizontalAlignment == false)
            {
                AddDeclaration("text-align", vAlign);
            }

            if (!string.IsNullOrEmpty(vAlign) && context.Exclude.VerticalAlignment == false)
            {
                AddDeclaration("vertical-align", vAlign);
            }

            return declarations;
        }

        protected string GetVerticalAlignment()
        {
            switch (_verticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    return "top";
                case ExcelVerticalAlignment.Center:
                    return "middle";
                case ExcelVerticalAlignment.Bottom:
                    return "bottom";
            }

            return "";
        }

        protected string GetHorizontalAlignment()
        {
            switch (_horizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    return "center";
                case ExcelHorizontalAlignment.Left:
                    return "left";
            }

            return "";
        }
    }
}
