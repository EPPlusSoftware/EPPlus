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

            if ((_horizontalAlignment != ExcelHorizontalAlignment.General) && (context.Exclude.HorizontalAlignment == false))
            {
                hAlign = GetHorizontalAlignment();
            }

            if ((_verticalAlignment != ExcelVerticalAlignment.Bottom) && (context.Exclude.VerticalAlignment == false))
            {
                vAlign = GetVerticalAlignment();
            }

            if (string.IsNullOrEmpty(hAlign) && _rightDefault)
            {
                hAlign = "right";
            }

            if (!string.IsNullOrEmpty(hAlign) && context.Exclude.HorizontalAlignment == false)
            {
                AddDeclaration("text-align", hAlign);
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
