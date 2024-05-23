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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils;

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
            var styles = _ws.Workbook.Styles;

            LowValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Min,
                ExcelConditionalFormattingConstants.Colors.CfvoLowValue, 
                priority, styles);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Max,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                priority, styles);
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
            var styles = _ws.Workbook.Styles;

            LowValue = new ExcelConditionalFormattingColorScaleValue(
            low,
            ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
            Priority, styles);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
            high,
            ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
            Priority, styles);

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
            var cellValue = ConvertUtil.GetValueDouble(range.Value);
            //TODO: Cache this for performance
            if (cellValue.IsNumeric())
            {
                var cellValues = new List<double>();
                foreach (var cell in Address.GetAllAddresses())
                {
                    for (int i = 1; i <= cell.Rows; i++)
                    {
                        for (int j = 1; j <= cell.Columns; j++)
                        {
                            cellValues.Add(ConvertUtil.GetValueDouble(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value));
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

                var newColor = CalculateNumberedGradient(index, values.Count()-1, LowValue.ColorSettings.GetColorAsColor(), HighValue.ColorSettings.GetColorAsColor());

                return "background-color:" + "#" + newColor.ToArgb().ToString("x8").Substring(2) + ";";
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
                    return true;
                }
            }

            return false;
        }

        double TruncateTo3Decimals(double value)
        {
            double ret = Math.Round(value * 100);
            return ret * 0.01;
        }

        protected Color LinearInterpolationTwoColors(Color color1, Color color2, double startPointWeight, double endPointWeight)
        {
            var newR = (int)Math.Round(color1.R * startPointWeight + color2.R * endPointWeight);
            var newG = (int)Math.Round(color1.G * startPointWeight + color2.G * endPointWeight);
            var newB = (int)Math.Round(color1.B * startPointWeight + color2.B * endPointWeight);

            return Color.FromArgb(1, newR, newG, newB);
        }

        protected Color CalculateNumberedGradient(double currentStep, double numStepsBetween, Color color1, Color color2)
        {
            double endPointWeight = currentStep / numStepsBetween;
            double startPointWeight = 1.0d - endPointWeight;

            //Lower accuracy to match excel
            startPointWeight = TruncateTo3Decimals(startPointWeight);
            endPointWeight = TruncateTo3Decimals(endPointWeight);

            return LinearInterpolationTwoColors(color1, color2, startPointWeight, endPointWeight);
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
                    ret = c._styles.GetIndexedColor(c.Index.Value);
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
