using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal static class FontScaleFactors
    {
        private static Dictionary<uint, FontScaleFactor> _fonts = new Dictionary<uint,FontScaleFactor>();
        private static object _syncRoot = new object();

        private static void Initialize()
        {
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Regular), new FontScaleFactor(1.13f, 1.05f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Bold), new FontScaleFactor(1.1f, 1.01f, 1f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Italic), new FontScaleFactor(1.1f, 1.03f, 1.02f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.BoldItalic), new FontScaleFactor(1.1f, 1.03f, 1.02f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Regular), new FontScaleFactor(1.1f, 1.05f, 1.04f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Bold), new FontScaleFactor(1.12f, 1.06f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Italic), new FontScaleFactor(1.12f, 1.09f, 1.06f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Arial, FontSubFamilies.BoldItalic), new FontScaleFactor(1.17f, 1.14f, 1.1f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Regular), new FontScaleFactor(1.09f, 1.02f, 1.02f, 1.3f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Bold), new FontScaleFactor(1.11f, 1.07f, 1.05f, 1.3f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Italic), new FontScaleFactor(1.09f, 1.01f, 1.01f, 1.3f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.BoldItalic), new FontScaleFactor(1.11f, 1.08f, 1.05f, 1.3f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Regular), new FontScaleFactor(1.12f, 1.06f, 1.02f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Bold), new FontScaleFactor(1.11f, 1.08f, 1f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Italic), new FontScaleFactor(1.09f, 1.09f, 1.02f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.BoldItalic), new FontScaleFactor(1.11f, 1.1f, 1.03f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Regular), new FontScaleFactor(1.11f, 1.03f, 1.02f, 1.2f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Bold), new FontScaleFactor(1.1f, 1.03f, 1.02f, 1.2f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Italic), new FontScaleFactor(1.1f, 1.03f, 1.02f, 1.2f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.BoldItalic), new FontScaleFactor(1.1f, 1.03f, 1.02f, 1.2f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Regular), new FontScaleFactor(1.09f, 1.03f, 1.02f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Bold), new FontScaleFactor(1.13f, 1.08f, 1.04f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Italic), new FontScaleFactor(1.13f, 1.08f, 1.07f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.BoldItalic), new FontScaleFactor(1.14f, 1.06f, 1.05f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Regular), new FontScaleFactor(1.17f, 1.1f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Bold), new FontScaleFactor(1.18f, 1.08f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Italic), new FontScaleFactor(1.16f, 1.08f, 1.07f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.BoldItalic), new FontScaleFactor(1.16f, 1.06f, 1.05f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Regular), new FontScaleFactor(1.13f, 1.07f, 1.07f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Bold), new FontScaleFactor(1.10f, 1.05f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Italic), new FontScaleFactor(1.10f, 1.04f, 1.04f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.BoldItalic), new FontScaleFactor(1.10f, 1.04f, 1.03f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Regular), new FontScaleFactor(1.13f, 1.04f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Bold), new FontScaleFactor(1.11f, 1.08f, 1.06f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Italic), new FontScaleFactor(1.08f, 1.04f, 1.04f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.BoldItalic), new FontScaleFactor(1.09f, 1.05f, 1.04f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Regular), new FontScaleFactor(1.13f, 1.07f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Bold), new FontScaleFactor(1.13f, 1.05f, 1.04f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Italic), new FontScaleFactor(1.13f, 1.07f, 1.06f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.BoldItalic), new FontScaleFactor(1.13f, 1.05f, 1.04f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Regular), new FontScaleFactor(1.09f, 1.04f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Bold), new FontScaleFactor(1.13f, 1.04f, 1.02f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Italic), new FontScaleFactor(1.09f, 1.04f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.BoldItalic), new FontScaleFactor(1.13f, 1.04f, 1.03f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Regular), new FontScaleFactor(1.12f, 1.03f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Bold), new FontScaleFactor(1.13f, 1.06f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Italic), new FontScaleFactor(1.13f, 1.06f, 1.05f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.BoldItalic), new FontScaleFactor(1.12f, 1.06f, 1.06f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Regular), new FontScaleFactor(1.12f, 1.07f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Bold), new FontScaleFactor(1.12f, 1.09f, 1.07f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Italic), new FontScaleFactor(1.13f, 1.15f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.BoldItalic), new FontScaleFactor(1.13f, 1.09f, 1.06f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Regular), new FontScaleFactor(1.1f, 1.05f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Bold), new FontScaleFactor(1.16f, 1.08f, 1.07f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Italic), new FontScaleFactor(1.1f, 1.04f, 1.03f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.BoldItalic), new FontScaleFactor(1.13f, 1.11f, 1.07f));

            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Regular), new FontScaleFactor(1.03f, 1.07f, 1.06f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Bold), new FontScaleFactor(1.11f, 1.02f, 1f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Italic), new FontScaleFactor(1.02f, 1.01f, 1f));
            _fonts.Add(GenericFontMetricsTextMeasurer.GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.BoldItalic), new FontScaleFactor(1.13f, 1.2f, 1.2f));
        }

        private static bool _initialized = false;

        public static float GetScaleFactor(uint key, float width)
        {
            if(!_initialized)
            {
                lock(_syncRoot)
                {
                    if (!_initialized)
                    {
                        _initialized = true;
                        Initialize();
                    }
                }
            }
            if(!_fonts.ContainsKey(key))
            {
                return 1f;
            }
            var factor = _fonts[key];
            return factor.Calculate(width);
        }
    }
}
