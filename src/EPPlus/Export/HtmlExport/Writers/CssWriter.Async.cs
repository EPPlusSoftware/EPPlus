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
using System;
using System.Linq;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
#if !NET35 && !NET40
    internal partial class CssWriter
    {
        internal async Task WritePropertyDeclarationAsync(Declaration declaration, bool minify)
        {
            await WriteCssItemAsync($"{declaration.Name}:{declaration.ValuesToString()};", minify);
        }

        internal async Task WriteRuleAsync(CssRule rule, bool minify)
        {
            await WriteSelectorOpeningAsync(rule.Selector, minify);

            for (int i = 0; i < rule.Declarations.Count; i++)
            {
                await WritePropertyDeclarationAsync(rule.Declarations[i], minify);
            }

            await WriteClassEndAsync(minify);
        }

        internal async Task WriteSelectorOpeningAsync(string selector, bool minify)
        {
            await WriteClassAsync($"{selector}{{", minify);
        }

        internal async Task WriteAndClearFlushAsync(CssRuleCollection collection, bool minify)
        {
			foreach (var rule in collection.CssRules.OrderByDescending(x => x.Order))
			{
				await WriteRuleAsync(rule, minify);
            }

            collection.CssRules.Clear();
            await _writer.FlushAsync();
        }
    }
#endif
}
