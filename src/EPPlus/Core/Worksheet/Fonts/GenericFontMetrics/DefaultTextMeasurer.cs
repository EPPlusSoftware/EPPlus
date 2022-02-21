/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal class DefaultTextMeasurer : GenericFontMetricsTextMeasurerBase
    {
        internal TextMeasurement Measure(string text, float size)
        {
            var fontKey = GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Regular);
            return MeasureTextInternal(text, fontKey, MeasurementFontStyles.Regular, size);
        }
    }
}
