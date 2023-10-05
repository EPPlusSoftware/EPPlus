using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Collectors
{

#if !NET35 && !NET40
    internal partial class CssRangeRuleCollection
    {
        private List<ExcelRangeBase> list;
        private HtmlRangeExportSettings settings;

        //public CssRangeRuleCollection(List<ExcelRangeBase> list, HtmlRangeExportSettings settings)
        //{
        //    this.list = list;
        //    this.settings = settings;
        //}

        //internal async Task AddSharedClassesAsync(string tableClass)
        //{
        //    if (_cssSettings.IncludeSharedClasses == false) return;

        //    await AddTableRule(tableClass);

        //    //Hidden class
        //    _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
        //    //Text-alignment classes
        //    _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
        //    _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");

        //    AddWorksheetDimensions();
        //    AddImageAlignment();
        //}

        internal async Task AddToCollectionAsync(List<ExcelXfs> xfsList, ExcelNamedStyleXml ns, int id)
        {
            var xfs = xfsList[0];

            var styleClass = new CssRule($".{settings.StyleClassPrefix}{settings.CellStyleClassName}{id}");
            var translators = new List<TranslatorBase>();

            if (xfs.FillId > 0)
            {
                translators.Add(new CssFillTranslator(xfs.Fill));
            }
            if (xfs.FontId > 0)
            {
                translators.Add(new CssFontTranslator(xfs.Font, ns.Style.Font));
            }

            if (xfsList.Count > 1)
            {
                var bXfs = xfsList[1];
                var rXfs = xfsList[2];

                if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                {
                    translators.Add(new CssBorderTranslator(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right));
                }
            }
            else if (xfs.BorderId > 0)
            {
                translators.Add(new CssBorderTranslator(xfs.Border));
            }

            translators.Add(new CssTextFormatTranslator(xfs));

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                await _context.AddDeclarationsAsync(styleClass);
            }

            _ruleCollection.CssRules.Add(styleClass);
        }
    }
#endif
}
