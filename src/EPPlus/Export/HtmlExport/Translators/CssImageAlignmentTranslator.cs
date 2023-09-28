using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssImageAlignmentTranslator : TranslatorBase
    {
        HtmlPictureSettings _picSettings;

        internal CssImageAlignmentTranslator(HtmlPictureSettings picSettings) 
        {
            _picSettings = picSettings;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            AddDeclaration("vertical-align", _picSettings.AddMarginTop ? "top" : "middle");
            AddDeclaration("text-align", _picSettings.AddMarginLeft ? "left" : "center");

            return declarations;
        }

    }
}
