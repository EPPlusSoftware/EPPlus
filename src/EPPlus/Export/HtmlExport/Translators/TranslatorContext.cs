using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        //internal ExcelXfs Xfs => _xfs;
        //internal ExcelNamedStyleXml Ns => _ns;
        //internal ExcelTheme Theme => _theme;

        //internal eBorderExclude BorderExclude => _borderExclude;
        //internal eFontExclude FontExclude => _fontExclude;

        private TranslatorBase strategy;

        public TranslatorContext(CssExclude exclude) 
        {
            Exclude = exclude;
        }

        public TranslatorContext(CssExclude exclude, TranslatorBase concreteStrategy)
        { 
            Exclude = exclude;
            strategy = concreteStrategy;
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

    }
}
