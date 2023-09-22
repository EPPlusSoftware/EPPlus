using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
    internal static class StyleToCss
    {
        private static bool IsAddedToCache(ExcelDxfStyleConditionalFormatting dxf, Dictionary<string, int> styleCache, out int id)
        {
            var key = dxf.Id;

            if (styleCache.ContainsKey(key))
            {
                id = styleCache[key];
                return true;
            }
            else
            {
                id = styleCache.Count + 1;
                styleCache.Add(key, id);
                return false;
            }
        }

        private static bool IsAddedToCache(ExcelXfs xfs, Dictionary<string, int> styleCache, out int id, int bottomStyleId = -1, int rightStyleId = -1)
        {
            var key = AttributeParser.GetStyleKey(xfs);
            if (bottomStyleId > -1) key += bottomStyleId + "|" + rightStyleId;
            if (styleCache.ContainsKey(key))
            {
                id = styleCache[key];
                return true;
            }
            else
            {
                id = styleCache.Count + 1;
                styleCache.Add(key, id);
                return false;
            }
        }


        internal static int GetIdFromCache(ExcelDxfStyleConditionalFormatting dxfs, Dictionary<string, int> styleCache)
        {
            if (dxfs != null)
            {
                if (!IsAddedToCache(dxfs, styleCache, out int id))
                {
                    return id;
                }
            }

            return -1;
        }

        //internal async Task StyleToCssStringAsync(ExcelDxfStyleConditionalFormatting dxfs, int id, string styleClassPrefix, string cellStyleClassName, EpplusCssWriter writer)
        //{
        //    var cls = $".{styleClassPrefix}{cellStyleClassName}-dxf-{id}{{";

        //    if (dxfs.Fill != null)
        //    {
        //        cls += await WriteFillStylesAsync(dxfs.Fill);
        //    }

        //    if (dxfs.Font != null)
        //    {
        //        await WriteFontStylesAsync(dxfs.Font);
        //    }

        //    if (dxfs.Border != null)
        //    {
        //        await WriteBorderStylesAsync(dxfs.Border.Top, dxfs.Border.Bottom, dxfs.Border.Left, dxfs.Border.Right);
        //    }

        //    await WriteClassEndAsync(_settings.Minify);

        //    return cls;
        //}



    }
}
