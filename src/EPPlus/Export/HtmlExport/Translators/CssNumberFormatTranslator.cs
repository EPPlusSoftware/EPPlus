using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssNumberFormatTranslator : TranslatorBase
    {
        string format;
        int id;
        ExcelIndexedColor? IndexColor = null;

        internal CssNumberFormatTranslator(INumberFormat numberFormat)
        {
            format = numberFormat.NumberFormatString;
            id = numberFormat.NumberFormatID;
            IndexColor = numberFormat.ColorId;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            var declarations = new List<Declaration>();
            if (IndexColor.HasValue)
            {
                var col2 = context.Theme._wb.Styles.GetIndexedColor((int)IndexColor.Value);
                var htmlColor = "#" + col2.ToArgb().ToString("x8").Substring(2);
                declarations.Add(new Declaration("color", htmlColor));
            }
            return declarations;
        }
    }
}
