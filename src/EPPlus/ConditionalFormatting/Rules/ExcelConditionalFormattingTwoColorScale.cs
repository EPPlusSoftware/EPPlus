/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Two Colour Scale class
    /// </summary>
    internal class ExcelConditionalFormattingTwoColorScale : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTwoColorScale
    {
        internal ExcelConditionalFormattingTwoColorScale(
        ExcelAddress address,
        int priority,
        ExcelWorksheet ws) 
        : base(
            eExcelConditionalFormattingRuleType.TwoColorScale, 
            address, 
            priority, 
            ws)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Min,
                ExcelConditionalFormattingConstants.Colors.CfvoLowValue, 
                priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Max,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                priority);
        }

        internal ExcelConditionalFormattingTwoColorScale(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet,
            bool stopIfTrue,
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal,
            string highVal,
            XmlReader xr) : base (
                eExcelConditionalFormattingRuleType.TwoColorScale, 
                address, 
                priority, 
                worksheet)
        {
            StopIfTrue = stopIfTrue;
            SetValues(low, high, lowVal, highVal);
            ReadColors(xr);
        }

        internal void SetValues(
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal = "",
            string highVal = "",
            string middleVal = "",
            eExcelConditionalFormattingValueObjectType? middle = null)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
            low,
            ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
            Priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
            high,
            ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
            Priority);

            if (!string.IsNullOrEmpty(lowVal))
            {
                if(double.TryParse(lowVal, out Double res))
                {
                    LowValue.Value = double.Parse(lowVal, CultureInfo.InvariantCulture);
                }
                else
                {
                    LowValue.Formula = lowVal;
                }
            }

            if (!string.IsNullOrEmpty(highVal))
            {
                if (double.TryParse(highVal, out Double res))
                {
                    HighValue.Value = double.Parse(highVal, CultureInfo.InvariantCulture);
                }
                else
                {
                    HighValue.Formula = highVal;
                }
            }
        }

        internal ExcelConditionalFormattingTwoColorScale(ExcelConditionalFormattingTwoColorScale copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
            LowValue = copy.LowValue;
            HighValue = copy.HighValue;
            StopIfTrue = copy.StopIfTrue;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingTwoColorScale(this, newWs);
        }

        internal override bool IsExtLst
        {
            get
            {
                if (ExcelAddressBase.RefersToOtherWorksheet(LowValue.Formula, _ws.Name) || 
                    ExcelAddressBase.RefersToOtherWorksheet(HighValue.Formula, _ws.Name))
                {
                    return true;
                }

                return base.IsExtLst;
            } 

        }

        internal virtual void ReadColors(XmlReader xr)
        {
            Type = eExcelConditionalFormattingRuleType.TwoColorScale;

            ReadColorAndColorSettings(xr, ref _lowValue);

            xr.Read();

            ReadColorAndColorSettings(xr, ref _highValue);

            xr.Read();
            xr.Read();
        }

        /// <summary>
        /// Internal Reading function
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="colSettings"></param>
        internal void ReadColorAndColorSettings(XmlReader xr, ref ExcelConditionalFormattingColorScaleValue colSettings)
        {
            if (!string.IsNullOrEmpty(xr.GetAttribute("auto")))
            {
                colSettings.ColorSettings.Auto = xr.GetAttribute("auto") == "1" ? true : false;
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("theme")))
            {
                colSettings.ColorSettings.Theme = (eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("indexed")))
            {
                colSettings.ColorSettings.Index = int.Parse(xr.GetAttribute("indexed"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("rgb")))
            {
                colSettings.ColorSettings.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
            {
                colSettings.ColorSettings.Tint = double.Parse(xr.GetAttribute("tint"), CultureInfo.InvariantCulture);
            }
        }

        internal ExcelConditionalFormattingColorScaleValue _lowValue;
        internal ExcelConditionalFormattingColorScaleValue _highValue;


        /// <summary>
        /// Low Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue LowValue
        {
            get { return _lowValue; }
            set { _lowValue = value; }
        }

        /// <summary>
        /// High Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue HighValue
        {
            get { return _highValue; }
            set { _highValue = value; }
        }

        internal virtual string ApplyStyleOverride(ExcelAddress address)
        {
            var range = _ws.Cells[address.Address];
            var cellValue = range.Value;
            if (cellValue.IsNumeric())
            {
                var cellValues = new List<object>();
                foreach (var cell in Address.GetAllAddresses())
                {
                    for (int i = 1; i <= cell.Rows; i++)
                    {
                        for (int j = 1; j <= cell.Columns; j++)
                        {
                            cellValues.Add(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value);
                        }
                    }
                }

                var values = cellValues.OrderBy(n => n);
                int index = 0;

                foreach(var value in values)
                {
                    if (value == cellValue)
                    {
                        break;
                    }
                    index++;
                }

                float percentage = (float)index / values.Count();

                var lowcol = LowValue.Color;
                var highcol = HighValue.Color;

                //var lerp = Vector3.Lerp(new Vector3(lowcol.R, lowcol.G,lowcol.B), new Vector3(highcol.R, highcol.G, highcol.B), percentage);

                var r = lowcol.R;
                var g = lowcol.G;
                var b = lowcol.B;

                var hiR = highcol.R;
                var hiG = highcol.G;
                var hiB = highcol.B;

                var originalPercent = 1.0 - percentage;
                //var absR = r + hiR * percentage;
                //var absG = g + hiG * percentage;
                //var absB = b + hiB * percentage;

                var absR = (int)Math.Abs(originalPercent * r - hiR * percentage);
                var absG = (int)Math.Abs(originalPercent * g - hiG * percentage);
                var absB = (int)Math.Abs(originalPercent * b - hiB * percentage);
                

                var newColor = Color.FromArgb(1, absR, absG, absB);

                //range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //range.Style.Fill.BackgroundColor.SetColor(newColor);

                return "#" + newColor.ToArgb().ToString("x8").Substring(2);

                //return "";
                //_ws._wb.ThemeManager.CurrentTheme;
            }
            return "";
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if (Address.Collide(address) != ExcelAddressBase.eAddressCollition.No)
            {
                var cellValue = _ws.Cells[address.Address].Value;

                if(cellValue.IsNumeric())
                {
                    //ApplyStyleOverride(address);
                    return true;
                    ////Formula2 only filled if there's a cell or formula to cond
                    //if (Formula2 != null)
                    //{
                    //    return _ws.Cells[Address.Start.Address].Formula.Contains(Formula2) ? false : true;
                    //}
                    //else
                    //{
                    //    return _ws.Cells[Address.Start.Address].Formula.Contains(_text) ? false : true;
                    //}
                }
            }

            return false;
        }

        protected string GetColor(ExcelDxfColor c, ExcelTheme theme)
        {
            Color ret;
            if (c.Color.HasValue)
            {
                ret = c.Color.Value;
            }
            else if (c.Theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, c.Theme.Value);
            }
            else if (c.Index != null)
            {
                if (c.Index.Value >= 0)
                {
                    ret = ExcelColor.GetIndexedColor(c.Index.Value);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }

            if (c.Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }

            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
