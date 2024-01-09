using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssDatabarTranslator : TranslatorBase
    {
        ExcelConditionalFormattingDataBar _databar;

        public CssDatabarTranslator(ExcelConditionalFormattingDataBar dataBar)
        {
            _databar = dataBar;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {

            return new List<Declaration> { };
        }


        internal void AddDatabar(bool isPositive, Color col)
        {
            //var ruleName = $".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}-databar-";
            //ruleName += isPositive ? $"positive-{id}" : $"negative-{id}";
            string turnDir = isPositive ? "0.25" : "0.75";

            var declarationVal = $"linear-gradient({turnDir}turn, rgba(0,{col.R},{col.G},{col.B}), 60%, white)";

            AddDeclaration("background-image", declarationVal);
            //var barClass = new CssRule(ruleName);

            //barClass.AddDeclaration("background-image", declarationVal);

            //_ruleCollection.CssRules.Add(barClass);
        }
    }
}
