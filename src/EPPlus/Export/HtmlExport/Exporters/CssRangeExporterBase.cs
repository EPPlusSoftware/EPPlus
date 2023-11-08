using OfficeOpenXml.Core;
using OfficeOpenXml.Export.HtmlExport.Collectors;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssRangeExporterBase : CssExporterBase
    {
        internal CssRangeExporterBase(HtmlExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges) 
        { }

        public CssRangeExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        { }

        protected CssRangeRuleCollection CreateRuleCollection(HtmlRangeExportSettings settings)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, settings);

            cssTranslator.AddSharedClasses(TableClass);

            AddRangesToCollection(cssTranslator);

            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(_ranges._list);
                foreach (var p in _rangePictures)
                {
                    cssTranslator.AddPictureToCss(p);
                }
            }

            return cssTranslator;
        }

        protected CssRangeRuleCollection CreateRuleCollection(HtmlTableExportSettings settings, ExcelTable table)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, settings);

            cssTranslator.AddSharedClasses(TableClass);

            if (settings.Css.IncludeTableStyles) cssTranslator.AddOtherCollectionToThisCollection
            (
                CreateTableCssRules(table, settings, _dataTypes).RuleCollection
            );

            if (settings.Css.IncludeCellStyles) AddCellCss(cssTranslator, table.Range, true);

            AddRangesToCollection(cssTranslator, true);


            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(_ranges._list);
                foreach (var p in _rangePictures)
                {
                    cssTranslator.AddPictureToCss(p);
                }
            }

            return cssTranslator;
        }
    }
}
