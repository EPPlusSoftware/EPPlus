/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Export.Pdf.Settings;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfDictionaries
    {
        internal readonly Dictionary<string, PdfFontResource> Fonts = new Dictionary<string, PdfFontResource>();
        internal readonly Dictionary<string, PdfPatternResource> Patterns = new Dictionary<string, PdfPatternResource>();
        internal readonly Dictionary<string, PdfShadingResource> Shadings = new Dictionary<string, PdfShadingResource>();
        internal Dictionary<string, IFontProvider> ShapedProviders = new Dictionary<string, IFontProvider>();

        public void AddFont(PdfPageSettings pageSettings, string FontName, FontSubFamily SubFamily, string Text)
        {
            var fullFont = FontName + " " + SubFamily.ToString();
            if (!Fonts.ContainsKey(fullFont))
            {
                int label = 1;
                if (Fonts.Count > 0)
                {
                    label = Fonts.Last().Value.labelNumber + 1;
                }
                Fonts.Add(fullFont, new PdfFontResource(FontName, SubFamily, label, pageSettings));
            }
            var manger = Fonts[fullFont].fontSubsetManager;
            manger.AddText(Text);
        }

        internal PdfFontResource GetFont(string FontName)
        {
            if (!Fonts.ContainsKey(FontName))
            {
                return null;
            }
            return Fonts[FontName];
        }

        //this should move and be on worksheet level.
        internal readonly Dictionary<string, PdfCommentsAndNotes> CommentsAndNotes = new Dictionary<string, PdfCommentsAndNotes>();
    }
}
