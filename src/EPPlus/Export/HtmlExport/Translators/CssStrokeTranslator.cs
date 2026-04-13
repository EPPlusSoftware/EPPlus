using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssStrokeTranslator : TranslatorBase
    {
        IDrawingBorder _drawingBorder;
        ExcelTheme _theme;

        public CssStrokeTranslator(IDrawingBorder drawingBorder) 
        {
            _drawingBorder = drawingBorder;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            var borderExclude = context.Exclude.Border;
            _theme = context.Theme;

            if(_drawingBorder.Stroke.HasValue && _drawingBorder.Stroke.PatternType == Style.ExcelFillStyle.Solid)
            {
                AddDeclaration($"stroke", _drawingBorder.Stroke.GetBackgroundColor(_theme));
            }
            return declarations;
        }
    }
}
