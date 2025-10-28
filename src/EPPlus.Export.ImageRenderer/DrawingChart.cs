/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusImageRenderer
{
    internal abstract class DrawingChart : DrawingBase
    {       
        public DrawingChart(ExcelChart chart) : base(chart) 
        {
            Chart = chart;
        }
        public ExcelChart Chart { get; set; }

        protected List<object> GetAxisValue(ExcelChartAxisStandard ax, RenderItem rect)
        {
            rect.GetBounds(out double x, out double y, out double w, out double h); 
            var values = ax.GetAxisValues();
            if(ax.AxisType==eAxisType.Cat)
            {
                return values.ToList();
            }
            var l = new List<object>();
            var min = double.MaxValue;
            var max = double.MinValue;
            foreach (var v in values)
            {
                var d = ConvertUtil.GetValueDouble(v, false, true);
                if(double.IsNaN(d))
                {
                    d = 0;
                }
                if(min>d)
                {
                    min = d;
                }
                if(max<d)
                {
                    max = d;
                }
            }
            var majorUnit = ax.MajorUnit ?? GetAutoUnit(min, max);

            return l;
        }

        private double GetAutoUnit(double min, double max)
        {
            var diff = max - min;

            var unit = diff / 10;
            var rest = unit % 10d;
            if (rest < 5)
            {
                unit -= rest;
            }
            else
            {
                unit += 10 - rest;
            }
            return unit;
        }
    }
}