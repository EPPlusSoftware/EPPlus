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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using OfficeOpenXml.ConditionalFormatting;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class TranslatorContext
    {

        internal ExcelTheme Theme;

        internal float IndentValue;
        internal string IndentUnit;

        internal CssExclude Exclude;
        internal CssExportSettings Settings;
        internal HtmlPictureSettings Pictures;

        private TranslatorBase strategy;

        internal bool SharedIconSetRuleAdded = false;
        internal bool SharedDatabarRulesAdded = false;

        internal HashSet<eExcelconditionalFormattingCustomIcon> AddedIcons = new HashSet<eExcelconditionalFormattingCustomIcon>();


        public TranslatorContext(HtmlRangeExportSettings settings) 
        {
            Exclude = settings.Css.CssExclude;
            Settings = settings.Css;
            Pictures = settings.Pictures;
        }

        public TranslatorContext(HtmlTableExportSettings settings, CssExclude exclude)
        {
            Settings = settings.Css;
            Pictures = settings.Pictures;
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
