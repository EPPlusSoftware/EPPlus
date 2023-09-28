using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssWorksheetDimensionsTranslator : TranslatorBase
    {
        List<ExcelWorksheet> _sheets;

        internal CssWorksheetDimensionsTranslator(List<ExcelWorksheet> sheets)
        {
            _sheets = sheets;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            throw new NotImplementedException();
        }


    }
}
