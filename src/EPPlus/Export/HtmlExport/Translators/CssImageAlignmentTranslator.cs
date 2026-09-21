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
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssImageAlignmentTranslator : TranslatorBase
    {
        HtmlDrawingSettings _drawingsSettings;

        internal CssImageAlignmentTranslator(HtmlDrawingSettings drawingsSettings) 
        {
            _drawingsSettings = drawingsSettings;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            AddDeclaration("vertical-align", _drawingsSettings.AddMarginTop ? "top" : "middle");
            AddDeclaration("text-align", _drawingsSettings.AddMarginLeft ? "left" : "center");

            return declarations;
        }

    }
}
