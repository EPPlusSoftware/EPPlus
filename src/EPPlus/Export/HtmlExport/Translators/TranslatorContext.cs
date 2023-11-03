using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class TranslatorContext
    {
        //ExcelXfs _xfs;
        //ExcelNamedStyleXml _ns;
        internal ExcelTheme Theme;
        //internal eBorderExclude BorderExclude;
        //internal eFontExclude FontExclude;
        //internal bool FillExclude;
        //internal bool WrapTextExclude;

        internal float IndentValue;
        internal string IndentUnit;

        internal CssExclude Exclude;
        internal CssExportSettings Settings;
        internal HtmlPictureSettings Pictures;

        //internal ExcelXfs Xfs => _xfs;
        //internal ExcelNamedStyleXml Ns => _ns;
        //internal ExcelTheme Theme => _theme;

        //internal eBorderExclude BorderExclude => _borderExclude;
        //internal eFontExclude FontExclude => _fontExclude;

        private TranslatorBase strategy;

        public TranslatorContext(HtmlRangeExportSettings settings) 
        {
            Exclude = settings.Css.CssExclude;
            Settings = settings.Css;
            Pictures = settings.Pictures;
        }

        public TranslatorContext(CssExclude exclude)
        { 
            Exclude = exclude;
        }

        public void SetTranslator(TranslatorBase concreteStrategy) 
        {
            strategy = concreteStrategy;
        }

        public void AddDeclarations(CssRule rule)
        {
            if(strategy == null)
            {
                throw new ArgumentNullException("Cannot add declarations without a Strategy! Try using .SetTranslator first");
            }

            rule.AddDeclarationList(strategy.GenerateDeclarationList(this));
        }

#if !NET35
        public async Task AddDeclarationsAsync(CssRule rule)
        {
            if (strategy == null)
            {
                throw new ArgumentNullException("Cannot add declarations without a Strategy! Try using .SetTranslator first");
            }

            rule.AddDeclarationList(await strategy.GenerateDeclarationListAsync(this));
        }
#endif

    }
}
