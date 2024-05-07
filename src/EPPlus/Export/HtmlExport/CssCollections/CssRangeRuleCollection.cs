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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Data.SqlTypes;

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
                _ruleCollection.AddRule($".{_settings.StyleClassPrefix}{_settings.IconPrefix}-shared::before", "content", "\"\"");

                var beforeRule = _ruleCollection.Last();
                //Set to 1.22 em because our standard-height for columns is 20px
                beforeRule.AddDeclaration("min-width", $"1.22em");
                beforeRule.AddDeclaration("min-height", $"1.22em");
                beforeRule.AddDeclaration("float", $"left");
                beforeRule.AddDeclaration("background-repeat", $"no-repeat");

                //Ensure cells don't overflow
                _ruleCollection.AddRule($".{_settings.StyleClassPrefix}{_settings.IconPrefix}-shared", "min-width", "2.24em");

                _context.SharedIconSetRuleAdded = true;
            }
        }

        internal void AddSharedDatabarRule(ExcelConditionalFormattingDataBar dataBar)
        {
            if(_context.SharedDatabarRulesAdded == false)
            {
                _ruleCollection.AddRule($".{_settings.StyleClassPrefix}{_settings.DatabarPrefix}-shared", "position", "relative");
                var sharedRule = _ruleCollection.Last();
                sharedRule.AddDeclaration("position", "relative");
                sharedRule.AddDeclaration("overflow", "hidden");
                sharedRule.AddDeclaration("background-image", $"url(data:image/svg+xml;base64,{DatabarSvg.GetConvertedAxisStripes()})");
                sharedRule.AddDeclaration("background-size", "5px 10px");
                sharedRule.AddDeclaration("background-repeat", "repeat-y");
                sharedRule.AddDeclaration("background-position", "-30px 0%");

                _ruleCollection.AddRule($".{_settings.StyleClassPrefix}{_settings.DatabarPrefix}-shared::after", "content", "\"\"");
                var sharedRuleAfter = _ruleCollection.Last();
                sharedRuleAfter.AddDeclaration("position", "absolute");
                sharedRuleAfter.AddDeclaration("width", "100%");
                sharedRuleAfter.AddDeclaration("height", "calc(100% - 3px)");
                sharedRuleAfter.AddDeclaration("z-index", "-1");
                sharedRuleAfter.AddDeclaration("top", "0%");
                sharedRuleAfter.AddDeclaration("bottom", "0%");
                sharedRuleAfter.AddDeclaration("background-repeat", "no-repeat");
                sharedRuleAfter.AddDeclaration("background-size", "100% 100%");

                _context.SharedDatabarRulesAdded = true;
            }
        }

        internal void AddDatabar(ExcelConditionalFormattingDataBar dataBar, int cssOrder, int id)
        {
            if (_context.SharedDatabarRulesAdded == false)
            {
                AddSharedDatabarRule(dataBar);
            }

            //TODO: Should cache and only create one of them if identical
            var ruleName = $".{_settings.StyleClassPrefix}{_settings.DxfStyleClassName}{id}-";
            var positiveDatabarRule = new CssRule(ruleName + "pos::after", cssOrder);
            var negativeDatabarRule = new CssRule(ruleName + "neg::after", cssOrder);

            var positiveDatabarSVG = DatabarSvg.GetConvertedDatabarString(dataBar.FillColor.GetColorAsColor(), dataBar.Gradient, dataBar.BorderColor.GetColorAsColor(true));
            var negativeDatabarSVG = DatabarSvg.GetConvertedDatabarString(dataBar.NegativeFillColor.GetColorAsColor(), dataBar.Gradient, dataBar.NegativeBorderColor.GetColorAsColor());

            positiveDatabarRule.AddDeclaration("background-image", $"url(data:image/svg+xml;base64,{positiveDatabarSVG})");
            negativeDatabarRule.AddDeclaration("background-image", $"url(data:image/svg+xml;base64,{negativeDatabarSVG})");

            _ruleCollection.AddRule($"{ruleName + "pos"}, {ruleName + "neg"}", "z-index", $"0");
            var sharedContentRule = _ruleCollection.Last();

            if (dataBar.AxisPosition != eExcelDatabarAxisPosition.None)
            {
                double AxisPositionPercent;
                if (dataBar.AxisPosition == eExcelDatabarAxisPosition.Automatic)
                {
                    var absLowest = Math.Abs(dataBar.lowest);
                    var absHighest = Math.Abs(dataBar.highest);

                    double denominator = dataBar.highest - dataBar.lowest;

                    var percent = Math.Abs(dataBar.lowest / denominator);
                    AxisPositionPercent = percent * 100;
                }
                else
                {
                    //is middle
                    AxisPositionPercent = 50;
                }

                if (dataBar.Direction == eDatabarDirection.RightToLeft)
                {
                    //Note: switches right/left and positive/negative
                    AxisPositionPercent = 100 - AxisPositionPercent;
                    var temp = positiveDatabarRule;
                    positiveDatabarRule = negativeDatabarRule;
                    negativeDatabarRule = temp;
                }

                AxisPositionPercent = Math.Round(AxisPositionPercent, 3);
                string xPositionInPercentString = (AxisPositionPercent).ToString(CultureInfo.InvariantCulture);

                if(dataBar.lowest < 0 && dataBar.highest > 0 | dataBar.AxisPosition == eExcelDatabarAxisPosition.Middle)
                {
                    sharedContentRule.AddDeclaration("background-position", $"{xPositionInPercentString}% 0%");
                }

                if (dataBar.AxisColor != null)
                {
                    var axisColorSvg = DatabarSvg.GetConvertedAxisStripesWithColor(dataBar.AxisColor.GetColorAsColor(true));
                    sharedContentRule.AddDeclaration("background-image", $"url(data:image/svg+xml;base64,{axisColorSvg})");
                }

                //Left of axis Negative, Right of axis Positive
                double negativeWidthPercent = AxisPositionPercent;
                double positiveWidthPercent = Math.Round(100d - AxisPositionPercent, 3);

                string negativeWidth = negativeWidthPercent.ToString(CultureInfo.InvariantCulture);
                string positiveWidth = positiveWidthPercent.ToString(CultureInfo.InvariantCulture);

                string rightOffset = positiveWidth;
                string leftOffset = negativeWidth;

                //Corrections for bar not to cover axis
                if (negativeWidthPercent > 50)
                {
                    negativeDatabarRule.AddDeclaration("background-position", $"2px");
                }
                else if(positiveWidthPercent < 50)
                {
                    positiveDatabarRule.AddDeclaration("background-position", $"2px");
                }
                else
                {
                    negativeDatabarRule.AddDeclaration("background-position", $"1px");
                    positiveDatabarRule.AddDeclaration("background-position", $"1px");
                }

                positiveDatabarRule.AddDeclaration("width", $"{positiveWidth}%");
                positiveDatabarRule.AddDeclaration("left", $"{leftOffset}%");

                negativeDatabarRule.AddDeclaration("width", $"{negativeWidth}%");
                negativeDatabarRule.AddDeclaration("right", $"{rightOffset}%");
                negativeDatabarRule.AddDeclaration("transform", $"scale(-1, 1)");
            }
            else
            {
                positiveDatabarRule.AddDeclaration("left", $"0%");
                negativeDatabarRule.AddDeclaration("left", $"0%");

                if (dataBar.Direction == eDatabarDirection.RightToLeft)
                {
                    positiveDatabarRule.AddDeclaration("transform", $"scale(-1, 1)");
                    negativeDatabarRule.AddDeclaration("transform", $"scale(-1, 1)");
                }
            }

            _ruleCollection.CssRules.Add(positiveDatabarRule);
            _ruleCollection.CssRules.Add(negativeDatabarRule);

            var addresses = dataBar.Address.GetAllAddresses();

            var cells = dataBar._ws.Cells[dataBar.Address.Address];

            foreach (var cell in cells)
            {
                var className = $".{_settings.StyleClassPrefix}{cell.Address}-{_settings.DatabarPrefix}::after";
                double percentage = dataBar.GetPercentageAtCell(cell);
                var databarRule = new CssRule(className, cssOrder);
                databarRule.AddDeclaration("background-size", $"{Math.Round(percentage, 3).ToString(CultureInfo.InvariantCulture)}% 100%");
                _ruleCollection.CssRules.Add(databarRule);
            }
        }

        internal void AddIconSetCF<T>(ExcelConditionalFormattingIconSetBase<T> set, int cssOrder, int id)
            where T : struct, Enum
        {
            if (_context.SharedIconSetRuleAdded == false)
            {
                AddSharedIconsetRule();
            }

            var ruleName = $".{_settings.StyleClassPrefix}{_settings.DxfStyleClassName}{id}";
            var contentRule = new CssRule(ruleName, cssOrder);
            if(!set.ShowValue)
            {
                contentRule.AddDeclaration("color", "transparent");
            }
            else
            {
                contentRule.AddDeclaration("visibility", "visible");
            }

            var icons = IconDict.GetIconsAsCustomIcons(set.GetIconSetString(), set.GetIconArray());

            for(int i = 0; i < icons.Length; i++)
            {
                if (_context.AddedIcons.Contains(icons[i]) == false)
                {
                    _context.AddedIcons.Add(icons[i]);
                    var svg = CF_Icons.GetIconSvg(icons[i]);

                    var iconRule = new CssRule($".{_settings.StyleClassPrefix}{_settings.IconPrefix}-{Enum.GetName(typeof(eExcelconditionalFormattingCustomIcon), icons[i])}::before", cssOrder);
                    iconRule.AddDeclaration("background-image", $" url(data:image/svg+xml;base64,{svg})");
                    _ruleCollection.CssRules.Add(iconRule);
                }
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
