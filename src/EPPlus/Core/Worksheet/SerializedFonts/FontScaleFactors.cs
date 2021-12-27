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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts
{
    internal class FontScaleFactors
    {
        private FontScaleFactors()
        {

        }

        private static Dictionary<uint, float> _factors;
        private static readonly object _syncRoot1 = new object();
        private static readonly object _syncRoot2 = new object();

        private static Dictionary<uint, float> CreateFactors()
        {
            var dict = new Dictionary<uint, float>();
            dict.Add(Key(SerializedFontFamilies.Arial, FontSubFamilies.Regular), 0.913f);
            dict.Add(Key(SerializedFontFamilies.Arial, FontSubFamilies.Bold), 0.975f);
            dict.Add(Key(SerializedFontFamilies.Arial, FontSubFamilies.Italic), 0.92f);
            dict.Add(Key(SerializedFontFamilies.Arial, FontSubFamilies.BoldItalic), 0.982f);

            dict.Add(Key(SerializedFontFamilies.Calibri, FontSubFamilies.Regular), 0.942f);
            dict.Add(Key(SerializedFontFamilies.Calibri, FontSubFamilies.Bold), 0.968f);
            dict.Add(Key(SerializedFontFamilies.Calibri, FontSubFamilies.Italic), 0.947f);
            dict.Add(Key(SerializedFontFamilies.Calibri, FontSubFamilies.BoldItalic), 0.975f);

            dict.Add(Key(SerializedFontFamilies.TimesNewRoman, FontSubFamilies.Regular), 0.916f);
            dict.Add(Key(SerializedFontFamilies.TimesNewRoman, FontSubFamilies.Bold), 0.975f);
            dict.Add(Key(SerializedFontFamilies.TimesNewRoman, FontSubFamilies.Italic), 0.907f);
            dict.Add(Key(SerializedFontFamilies.TimesNewRoman, FontSubFamilies.BoldItalic), 0.945f);

            dict.Add(Key(SerializedFontFamilies.CourierNew, FontSubFamilies.Regular), 0.972f);
            dict.Add(Key(SerializedFontFamilies.CourierNew, FontSubFamilies.Bold), 0.972f);
            dict.Add(Key(SerializedFontFamilies.CourierNew, FontSubFamilies.Italic), 0.972f);
            dict.Add(Key(SerializedFontFamilies.CourierNew, FontSubFamilies.BoldItalic), 0.972f);

            dict.Add(Key(SerializedFontFamilies.LiberationSerif, FontSubFamilies.Regular), 0.918f);
            dict.Add(Key(SerializedFontFamilies.LiberationSerif, FontSubFamilies.Bold), 0.978f);
            dict.Add(Key(SerializedFontFamilies.LiberationSerif, FontSubFamilies.Italic), 0.911f);
            dict.Add(Key(SerializedFontFamilies.LiberationSerif, FontSubFamilies.BoldItalic), 0.948f);
            return dict;
        }

        private static uint Key(SerializedFontFamilies f, FontSubFamilies s)
        {
            return SerializedFontMetrics.GetKey(f, s);
        }

        public static Dictionary<uint, float> Instance
        {
            get
            {
                lock(_syncRoot1)
                {
                    if (_factors == null)
                    {
                        lock (_syncRoot2)
                        {
                            _factors = CreateFactors();
                        }
                    }
                }
                return _factors; 
            }
        }
    }
}
