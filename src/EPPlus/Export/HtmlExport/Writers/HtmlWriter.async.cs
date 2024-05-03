/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class HtmlWriter
    {
#if !NET35 && !NET40
        public async Task RenderBeginTagAsync(string elementName, List<EpplusHtmlAttribute> attributes = null, bool closeElement = false)
        {
            _newLine = false;
            if (elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                await WriteIndentAsync();
            }
            await _writer.WriteAsync($"<{elementName}");

            if (attributes != null)
            {
                foreach (var attribute in attributes)
                {
                    await _writer.WriteAsync($" {attribute.AttributeName}=\"{attribute.Value}\"");
                }
                attributes.Clear();
            }

            if (closeElement)
            {
                await _writer.WriteAsync("/>");
                await _writer.FlushAsync();
            }
            else
            {
                await _writer.WriteAsync(">");
            }
        }

        public async Task RenderEndTagAsync(string elementName)
        {
            if (_newLine)
            {
                await WriteIndentAsync();
            }

            await _writer.WriteAsync($"</{elementName}>");
            await _writer.FlushAsync();
        }

        public async Task RenderHTMLElementAsync(HTMLElement element, bool minify)
        {
            await RenderBeginTagAsync(element.ElementName, element._attributes, element.IsVoidElement);

            if (element.IsVoidElement)
            {
                if (element.ElementName != HtmlElements.Img)
                {
                    await ApplyFormatAsync(minify);
                }
                return;
            }

            if (element._childElements.Count > 0)
            {
                var name = element.ElementName;
                bool noIndent = minify == true ? true : HtmlElements.NoIndentElements.Contains(name);

                await ApplyFormatIncreaseIndentAsync(noIndent);

                foreach (var child in element._childElements)
                {
                    await RenderHTMLElementAsync(child, minify);
                }

                if (noIndent == false)
                {
                    Indent--;
                }
            }

            await WriteAsync(element.Content);

            await RenderEndTagAsync(element.ElementName);
            if (element.ElementName != "a")
            {
                await ApplyFormatAsync(minify);
            }
        }
#endif
    }
}
