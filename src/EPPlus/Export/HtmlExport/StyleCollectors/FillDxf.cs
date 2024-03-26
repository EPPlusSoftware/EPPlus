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
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System.Drawing;
using System;
using OfficeOpenXml.Drawing.Theme;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FillDxf : IFill
    {
        ExcelDxfFill _fill;

        public FillDxf(ExcelDxfFill fill)
        {
            _fill = fill;
        }

        public ExcelFillStyle PatternType 
        { 
            get 
            {
                if (_fill.HasValue)
                {
                    if(_fill.PatternType.HasValue)
                    {
						return _fill.PatternType.Value;						
                    }
                    return ExcelFillStyle.Solid; 
				}
                return ExcelFillStyle.None;
            } 
        }

        public bool IsGradient
        {
            get
            {
                return _fill.Gradient != null;
            }
        }

        public double Degree
        {
            get
            {
                if (IsGradient && _fill.Gradient.Degree.HasValue)
                {
                    return _fill.Gradient.Degree.Value;
                }

                return double.NaN;
            }
        }

        public double Right
        {
            get
            {
                if (IsGradient && _fill.Gradient.Right.HasValue)
                {
                    return _fill.Gradient.Right.Value;
                }

                return double.NaN;
            }
        }

        public double Bottom
        {
            get
            {
                if (IsGradient && _fill.Gradient.Bottom.HasValue)
                {
                    return _fill.Gradient.Bottom.Value;
                }

                return double.NaN;
            }
        }

        public bool IsLinear
        {
            get
            {
                return _fill.Gradient.GradientType == eDxfGradientFillType.Linear;
            }
        }

        public string GetBackgroundColor(ExcelTheme theme)
        {
			return GetColor(_fill.BackgroundColor, theme);
		}

        public string GetPatternColor(ExcelTheme theme)
        {
			return GetColor(_fill.PatternColor, theme);			
        }

        public string GetGradientColor1(ExcelTheme theme)
        {
            return GetColor(_fill.Gradient.Colors[0].Color, theme);
        }
        public string GetGradientColor2(ExcelTheme theme)
        {
            return GetColor(_fill.Gradient.Colors[1].Color, theme);
        }

        public bool HasValue
        {
            get
            {
                return _fill.HasValue;
            }
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
                    ret = theme._wb.Styles.GetIndexedColor(c.Index.Value);
                }
                else 
                {
                    ret = Color.Empty;
                }
            }
            else
            {
                //Automatic, set to black.
                if (c.Auto.HasValue && c.Auto.Value)
                {
                    ret = Color.Empty;
                }
                else
                {
                    return null;
                }
            }

            if (c.HasValue && c.Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }

            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
