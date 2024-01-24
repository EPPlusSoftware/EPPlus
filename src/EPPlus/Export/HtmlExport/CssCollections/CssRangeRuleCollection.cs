using OfficeOpenXml.ConditionalFormatting;
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

        bool _hasAddedDBGenerics = false;

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
            _hasAddedDBGenerics = false;
        }

        internal void AddSharedClasses(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;

            AddTableRule(tableClass);

            //Hidden class
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
            //Text-alignment classes
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");

            AddWorksheetDimensions();
            AddImageAlignment();
        }

        internal void AddDatabarSharedClasses()
        {
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}pRelParent", "position", "relative; width: 100%; height: 100%"); //Relative position parent
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}relChildLeft", "float", "left; height: 100%"); //LeftChild
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}relChildRight", "overflow", "hidden; height: 100%"); //Right child

            //Border classes
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}dbr ", "border-right", "dashed"); //Dashed Border Right
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}dbl ", "border-left", "dashed"); //Dashed Border Left

            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}dbc", "width", "100%; height: 100%; position: absolute; display: flex"); //databarcontent
        }

        private void AddTableRule(string tableClass)
        {
            var tableRule = new CssRule($"table.{tableClass}");

            _context.SetTranslator(new CssTableTranslator(_wb.Styles.GetNormalStyle()));
            _context.AddDeclarations(tableRule);
            _ruleCollection.AddRule(tableRule);
        }

        private void AddImageAlignment()
        {
            if (_settings.Pictures.Include != ePictureInclude.Exclude && _settings.Pictures.CssExclude.Alignment == false)
            {
                var imgClass = new CssRule($"td.{_settings.StyleClassPrefix}image-cell ");
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
                CssRule widthRule = new CssRule(clsName);
                widthRule.AddDeclaration("width", $"{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px");

                clsName = "." + HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1) + " ";
                CssRule heightRule = new CssRule(clsName);
                heightRule.AddDeclaration("height", $"{(int)(ws.DefaultRowHeight / 0.75)}px");

                _ruleCollection.AddRule(widthRule);
                _ruleCollection.AddRule(heightRule);
            }
        }

        internal void AddDatabar(ExcelConditionalFormattingDataBar bar)
        {
            AddDatabar(bar, true);
            AddDatabar(bar, false);

            if(bar.AxisPosition == eExcelDatabarAxisPosition.Automatic)
            {
                var res = Math.Abs(bar.highest) + Math.Abs(bar.lowest);
                var axisPercent = bar.lowest < 0 && bar.highest > 0 ? Math.Abs(bar.lowest) / Math.Abs(bar.highest) : 0;

                var resFinal = (axisPercent * 100) >= 100 ? 90 : 0;

                var barClass = new CssRule($".leftWidth{bar.DxfId}");
                barClass.AddDeclaration("width", $" {(resFinal).ToString(CultureInfo.InvariantCulture)}%");
                _ruleCollection.CssRules.Add(barClass);
            }

            if (_hasAddedDBGenerics == false)
            {
                AddDatabarSharedClasses();
                AddDatabarGeneric(".pos-dbar", true);
                AddDatabarGeneric(".neg-dbar", false);
                _hasAddedDBGenerics = true;
            }
        }

        internal void AddDatabarGeneric(string ruleName, bool isPositive)
        {
            var barClass = new CssRule(ruleName);

            barClass.AddDeclaration("background-repeat", "no-repeat");

            if (isPositive)
            {
                barClass.AddDeclaration("background-position", "center left");
            }
            else
            {
                barClass.AddDeclaration("background-position", "center right");
            }

            _ruleCollection.CssRules.Add(barClass);
        }


        internal void AddDatabar(ExcelConditionalFormattingDataBar bar, bool isPositive = true)
        {
            var id = bar.DxfId;

            var ruleName = $".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}-db-";

            ruleName += isPositive ? $"pos{id}" : $"neg{id}";

            var firstTest = _ruleCollection.CssRules.Where(db => db.Selector == ruleName).Select(db => db);
            var test = firstTest.FirstOrDefault();
            if (test == null)
            {
                var barClass = new CssRule(ruleName);
                string turnDir = "0.25";
                Color col = Color.Empty;
                Color borderColor = Color.Empty;

                if (isPositive) 
                { 
                    col = bar.FillColor.GetColorAsColor();
                    borderColor = bar.BorderColor.GetColorAsColor();

                    if (bar.AxisPosition != eExcelDatabarAxisPosition.None && bar.lowest < 0)
                    {
                        barClass.AddDeclaration("border-left", "dotted");

                        var aColor = bar.AxisColor.GetColorAsColor();
                        if (aColor == Color.Empty)
                        {
                            aColor = Color.Black;
                        }
                        barClass.AddDeclaration($"border-left-color", $"rgb({aColor.R}, {aColor.G}, {aColor.B})");
                    }
                }
                else
                {
                    turnDir = "0.75";
                    col = bar.NegativeFillColor.GetColorAsColor();
                    borderColor = bar.NegativeBorderColor.GetColorAsColor();

                    if (bar.AxisPosition != eExcelDatabarAxisPosition.None && bar.highest > 0)
                    {
                        barClass.AddDeclaration("border-right", "dotted");

                        var aColor = bar.AxisColor.GetColorAsColor();
                        if (aColor == Color.Empty)
                        {
                            aColor = Color.Black;
                        }
                        barClass.AddDeclaration($"border-right-color",$"rgb({aColor.R}, {aColor.G}, {aColor.B})");
                    }
                }

                var declarationVal = $"linear-gradient({turnDir}turn, rgb({col.R},{col.G},{col.B}), 60%, white)";
                if (bar.Border)
                {
                    var borderRGB = $"{borderColor.R},{borderColor.G},{borderColor.B}";
                    declarationVal += $", linear-gradient({turnDir}turn, rgb({borderRGB}), 60%, rgb({borderRGB}))";
                }
                barClass.AddDeclaration("background-image", declarationVal);

                //barClass.AddDeclaration("border-color", barClass.AxisColor);
                //barClass.AddDeclaration("image-border-color", barClass.);

                _ruleCollection.CssRules.Add(barClass);
            }
        }

        internal void AddToCollection(List<IStyleExport> styleList, ExcelNamedStyleXml ns, int id, string altName = null)
        {
            var style = styleList[0];
            var ruleName = altName == null ? $".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}{id}" : altName;

            var styleClass = new CssRule(ruleName);
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

                if (style.Border != null && style.Border.HasValue || bXfs.Border != null && bXfs.Border.HasValue || rXfs.Border != null && rXfs.Border.HasValue)
                {
                    translators.Add(new CssBorderTranslator(style.Border.Top, bXfs.Border.Bottom, style.Border.Left, rXfs.Border.Right));
                }
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

        internal void AddPictureToCss(HtmlImage p)
        {
            var translator = new CssImageTranslator(p);

            if (translator.type == null) return;

            var pc = (IPictureContainer)p.Picture;

            if (_images.Contains(pc.ImageHash) == false)
            {
                string imageFileName = HtmlExportImageUtil.GetPictureName(p);
                var imgRule = new CssRule($"img.{_settings.StyleClassPrefix}image-{imageFileName}");

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

            var imgProperties = new CssRule($"img.{_settings.StyleClassPrefix}image-prop-{imageName}");
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
