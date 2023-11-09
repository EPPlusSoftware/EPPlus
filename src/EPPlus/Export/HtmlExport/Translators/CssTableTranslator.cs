using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssTableTranslator : TranslatorBase
    {
        ExcelNamedStyleXml _ns;

        public CssTableTranslator(ExcelNamedStyleXml ns) 
        {
            _ns = ns;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            if (context.Settings.IncludeSharedClasses == false) return null;

            if (context.Settings.IncludeNormalFont)
            {
                if (_ns != null)
                {
                    AddDeclaration($"font-family", _ns.Style.Font.Name);
                    AddDeclaration($"font-size", $"{_ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt");
                }
            }

            foreach (var item in context.Settings.AdditionalCssElements)
            {
                AddDeclaration(item.Key, item.Value);
            }

            return declarations;
        }
    }
}
