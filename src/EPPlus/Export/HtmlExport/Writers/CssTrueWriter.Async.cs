using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
#if !NET35 && !NET40
    internal partial class CssTrueWriter
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
    }
#endif
}
