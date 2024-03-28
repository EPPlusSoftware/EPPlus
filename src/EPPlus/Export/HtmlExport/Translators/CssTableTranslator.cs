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
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.Globalization;

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
                    //AddDeclaration($"height","100%");
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
