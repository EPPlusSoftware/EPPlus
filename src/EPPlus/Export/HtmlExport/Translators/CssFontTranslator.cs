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

using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.CssCollections;

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
                declarations.Add(new Declaration("color", _f.Color.GetColor(context.Theme)));
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
