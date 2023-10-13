using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;
using System.Xml.Linq;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssFontTranslator : TranslatorBase
    {
        IFont _f;
        ExcelFont _nf;

        internal CssFontTranslator(IFont f, ExcelFont nf) : base() 
        {
            _f = f;
            _nf = nf;
        }


        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            //ExcelFontXml f, FontDeclarationRules rules, ExcelFont nf, eFontExclude fontExclude, ExcelTheme theme

            var declarations = new List<Declaration>();
            var fontExclude = context.Exclude.Font;
            var fontRules = new FontDeclarationRules(_f, _nf, context);

            if (fontRules.HasFamily)
            {
                declarations.Add(new Declaration("font-family", _f.Name));
            }
            if (fontRules.HasSize)
            {
                declarations.Add(new Declaration("font-size", $"{_f.Size.ToString("g", CultureInfo.InvariantCulture)}pt"));
            }
            if (fontRules.HasColor)
            {
                declarations.Add(new Declaration("color", HtmlUtils.ColorUtils.GetColor(_f.Color, context.Theme)));
            }
            if (fontRules.HasBold)
            {
                declarations.Add(new Declaration("font-weight", "bolder"));
            }
            if (fontRules.HasItalic)
            {
                declarations.Add(new Declaration("font-style", "italic"));
            }
            if (fontRules.HasStrike)
            {
                declarations.Add(new Declaration("text-decoration", "line-through", "solid"));
            }
            if (fontRules.HasUnderline)
            {
                switch (_f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        declarations.Add(new Declaration("text-decoration", "underline", "double"));
                        break;
                    default:
                        declarations.Add(new Declaration("text-decoration", "underline", "solid"));
                        break;
                }
            }

            return declarations;
        }
    }
}
