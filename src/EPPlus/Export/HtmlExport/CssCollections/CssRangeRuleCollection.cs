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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
{
    internal partial class CssRangeRuleCollection
    {
        HtmlExportSettings _settings;
        CssExportSettings _cssSettings;

        ExcelWorkbook _wb;
        List<ExcelRangeBase> _ranges;
        ExcelTheme _theme;

        CssRuleCollection _ruleCollection;

        internal protected HashSet<string> _images = new HashSet<string>();

        internal CssRuleCollection RuleCollection => _ruleCollection;

        TranslatorContext _context;

        internal CssRangeRuleCollection(List<ExcelRangeBase> ranges, HtmlRangeExportSettings settings)
        {
            _settings = settings;
            _cssSettings = settings.Css;
            Init(ranges);
            _ruleCollection = new CssRuleCollection();

            _context = new TranslatorContext(settings);
            _context.Theme = _theme;
            _context.IndentValue = _cssSettings.IndentValue;
            _context.IndentUnit = _cssSettings.IndentUnit;
        }

        internal CssRangeRuleCollection(List<ExcelRangeBase> ranges, HtmlTableExportSettings settings)
        {
            _settings = settings;
            _cssSettings = settings.Css;
            Init(ranges);
            _ruleCollection = new CssRuleCollection();

            _context = new TranslatorContext(settings, settings.Css.Exclude.TableStyle);
            _context.Theme = _theme;
            _context.IndentValue = _cssSettings.IndentValue;
            _context.IndentUnit = _cssSettings.IndentUnit;
        }

        private void Init(List<ExcelRangeBase> ranges)
        {
            _ranges = ranges;
            _wb = _ranges[0].Worksheet.Workbook;
            if (_wb.ThemeManager.CurrentTheme == null)
            {
                _wb.ThemeManager.CreateDefaultTheme();
            }
            _theme = _wb.ThemeManager.CurrentTheme;
        }

        internal void AddSharedClasses(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;

            //Css reset. Different browsers have different default values.
            if(_cssSettings.IncludeCssReset)
            {
                _ruleCollection.AddRule("* ", "margin", "0; padding:0");
            }

            AddTableRule(tableClass);

            //Hidden class
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
            //Text-alignment classes
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");

            AddWorksheetDimensions();
            AddImageAlignment();
        }

        private void AddTableRule(string tableClass)
        {
            var tableRule = new CssRule($"table.{tableClass}", int.MaxValue);

            _context.SetTranslator(new CssTableTranslator(_wb.Styles.GetNormalStyle()));
            _context.AddDeclarations(tableRule);
            _ruleCollection.AddRule(tableRule);
        }

        private void AddImageAlignment()
        {
            if (_settings.Pictures.Include != ePictureInclude.Exclude && _settings.Pictures.CssExclude.Alignment == false)
            {
                var imgClass = new CssRule($"td.{_settings.StyleClassPrefix}image-cell ", int.MaxValue);
                imgClass.AddDeclaration("vertical-align", _settings.Pictures.AddMarginTop ? "top" : "middle");
                imgClass.AddDeclaration("text-align", _settings.Pictures.AddMarginLeft ? "left" : "center");

                _ruleCollection.AddRule(imgClass);
            }
        }

        private void AddWorksheetDimensions()
        {
            var worksheets = _ranges.Select(x => x.Worksheet).Distinct().ToList();
            foreach (var ws in worksheets)
            {
                var clsName = "." + HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1) + " ";
                CssRule widthRule = new CssRule(clsName, int.MaxValue);
                widthRule.AddDeclaration("width", $"{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px");

                clsName = "." + HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1) + " ";
                CssRule heightRule = new CssRule(clsName, int.MaxValue);
                heightRule.AddDeclaration("height", $"{(int)(ws.DefaultRowHeight / 0.75)}px");

                _ruleCollection.AddRule(widthRule);
                _ruleCollection.AddRule(heightRule);
            }
        }

        internal void AddToCollection(List<IStyleExport> styleList, ExcelNamedStyleXml ns, int id, int cssOrder, string altName = null)
        {
            var style = styleList[0];
            var ruleName = altName == null ? $".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}{id}" : altName;

            var styleClass = new CssRule(ruleName, cssOrder);
            var translators = new List<TranslatorBase>();

            if (style.Fill != null && style.Fill.HasValue && _context.Exclude.Fill == false)
            {
                translators.Add(new CssFillTranslator(style.Fill));
            }
            if (style.Font != null && style.Font.HasValue && _context.Exclude.Font != eFontExclude.All)
            {
                translators.Add(new CssFontTranslator(style.Font, ns.Style.Font));
            }

            if (styleList.Count > 1)
            {
                var bXfs = styleList[1];
                var rXfs = styleList[2];

                IBorder topLeft = style.Border ?? null;
                IBorder bottom = bXfs.Border ?? null;
                IBorder right = rXfs.Border ?? null;

                var borderTranslator = new CssBorderTranslator(topLeft, bottom, right);
                translators.Add(borderTranslator);
            }
            else if (style.Border != null && style.Border.HasValue)
            {
                translators.Add(new CssBorderTranslator(style.Border));
            }

            if (style is StyleXml)
            {
                translators.Add(new CssTextFormatTranslator((StyleXml)style));
            }

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            _ruleCollection.CssRules.Add(styleClass);
        }

        internal void AddSharedIconsetRule()
        {
            if(_context.SharedIconSetRuleAdded == false)
            {
                _ruleCollection.AddRule(".cf-ic-shared::before", "content", "\"\"");

                var beforeRule = _ruleCollection.Last();
                //TODO: Figure out where the .22 we need to account for comes from.
                //We add this because 1em does not make the svg box the same size as text of the font otherwise.
                //Likely due to some kind of margin setting somewhere?
                beforeRule.AddDeclaration("min-width", $"1.22em");
                beforeRule.AddDeclaration("min-height", $"1.22em");
                beforeRule.AddDeclaration("float", $"left");
                beforeRule.AddDeclaration("background-repeat", $"no-repeat");

                //Ensure cells don't overflow
                _ruleCollection.AddRule(".cf-ic-shared", "min-width", "2.24em");

                _context.SharedIconSetRuleAdded = true;
            }
        }

        internal void AddAdvancedCF<T>(ExcelConditionalFormattingIconSetBase<T> set, int cssOrder, int id)
            where T : struct, Enum
        {
            if (_context.SharedIconSetRuleAdded == false)
            {
                AddSharedIconsetRule();
            }

            var ruleName = $".{_settings.StyleClassPrefix}{_settings.DxfStyleClassName}cf{id}";
            var contentRule = new CssRule(ruleName, cssOrder);
            if(!set.ShowValue)
            {
                contentRule.AddDeclaration("color", "transparent");
            }
            else
            {
                contentRule.AddDeclaration("visibility", "visible");
            }

            string[] svgs;

            if (set.Custom)
            {
                svgs = CF_Icons.GetIconSetSvgsWithCustoms(set.GetIconSetString(), set.GetIconArray());
            }
            else
            {
                svgs = CF_Icons.GetIconSetSvgs(set.GetIconSetString());
            }

            for (int i = 0; i < svgs.Length; i++)
            {
                var iconRule = new CssRule(ruleName + $"{i}::before", cssOrder);
                iconRule.AddDeclaration("background-image", $" url(data:image/svg+xml;base64,{svgs[i]})");
                _ruleCollection.CssRules.Add(iconRule);
            }

            _ruleCollection.CssRules.Add(contentRule);
        }



        internal void AddPictureToCss(HtmlImage p)
        {
            var translator = new CssImageTranslator(p);

            if (translator.type == null) return;

            var pc = (IPictureContainer)p.Picture;

            if (_images.Contains(pc.ImageHash) == false)
            {
                string imageFileName = HtmlExportImageUtil.GetPictureName(p);
                var imgRule = new CssRule($"img.{_settings.StyleClassPrefix}image-{imageFileName}", int.MaxValue);

                _context.SetTranslator(translator);
                _context.AddDeclarations(imgRule);

                _ruleCollection.AddRule(imgRule);
                _images.Add(pc.ImageHash);
            }
            AddPicturePropertiesToCss(p);
        }

        private void AddPicturePropertiesToCss(HtmlImage image)
        {
            string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);

            var imgProperties = new CssRule($"img.{_settings.StyleClassPrefix}image-prop-{imageName}", int.MaxValue);
            _context.SetTranslator(new CssImagePropertiesTranslator(image));
            _context.AddDeclarations(imgProperties);

            RuleCollection.AddRule(imgProperties);
        }

        internal void AddOtherCollectionToThisCollection(CssRuleCollection otherCollection)
        {
            foreach (var otherRule in otherCollection)
            {
                _ruleCollection.AddRule(otherRule);
            }
        }
    }
}
