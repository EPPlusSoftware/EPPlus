/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting.Processors;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetFontBuilder
    {

        private static IEnumerable<IFontSubsetProcessor> Processors => new List<IFontSubsetProcessor>
        {
            new GlyfAndLocaSubsetProcessor(),
            new HeadSubsetProcessor(),
            new NameSubsetProcessor(),
            new MaxpSubsetProcessor(),
            new HheaSubsetProcessor(),
            new HmtxSubsetProcessor(),
            new CmapSubsetProcessor(),
            new Os2SubsetProcessor(),
            new PostSubsetProcessor(),
            new KernSubsetProcessor()
        };

        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            var context = new FontSubsettingContext(originalFont, unicodeChars);

            foreach(var processor in Processors)
            {
                processor.Process(context);
            }

            // 9. Debug-info
            context.SubsetFont.UsedCodePointsForSubset = new List<uint>(context.UsedCodePoints);

            return context.SubsetFont;
        }
    }
}