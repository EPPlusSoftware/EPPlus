using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal class FontScaleFactors
    {
        private Dictionary<uint, FontScaleFactor> _fonts = new Dictionary<uint,FontScaleFactor>();

        public FontScaleFactors()
        {
            Initialize();
        }

        public static readonly float JapaneseKanjiDefaultScalingFactor = 1.03f;

        private static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            return GenericFontMetricsTextMeasurer.GetKey(family, subFamily);
        }

        private static FontScaleFactor CSF(float s, float m, float l)
        {
            return new FontScaleFactor(s, m, l);
        }

        private static FontScaleFactor CSF(float s, float m, float l, float sf)
        {
            return new FontScaleFactor(s, m, l, sf);
        }
        private void Initialize()
        {
            _fonts.Add(GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Regular), CSF(1.33f, 1.09f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Bold), CSF(1.34f, 1.1f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.Italic), CSF(1.3f, 1.12f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.Calibri, FontSubFamilies.BoldItalic), CSF(1.3f, 1.1f, 1.03f));

            _fonts.Add(GetKey(FontMetricsFamilies.CalibriLight, FontSubFamilies.Regular), CSF(1.21f, 1.12f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalibriLight, FontSubFamilies.Bold), CSF(1.22f, 1.1f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalibriLight, FontSubFamilies.Italic), CSF(1.25f, 1.1f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalibriLight, FontSubFamilies.BoldItalic), CSF(1.22f, 1.09f, 1.03f));

            _fonts.Add(GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Regular), CSF(1.15f, 1.06f, 1.02f));
            _fonts.Add(GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Bold), CSF(1.17f, 1.12f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.Arial, FontSubFamilies.Italic), CSF(1.13f, 1.09f, 1.04f));
            _fonts.Add(GetKey(FontMetricsFamilies.Arial, FontSubFamilies.BoldItalic), CSF(1.18f, 1.12f, 1.07f));

            _fonts.Add(GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Regular), CSF(1.24f, 1.10f, 1.05f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Bold), CSF(1.26f, 1.10f, 1.08f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.Italic), CSF(1.25f, 1.05f, 1.06f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialBlack, FontSubFamilies.BoldItalic), CSF(1.20f, 1.17f, 1.07f, 1.3f));

            _fonts.Add(GetKey(FontMetricsFamilies.ArialNarrow, FontSubFamilies.Regular), CSF(1.15f, 1.18f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialNarrow, FontSubFamilies.Bold), CSF(1.23f, 1.24f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialNarrow, FontSubFamilies.Italic), CSF(1.25f, 1.14f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.ArialNarrow, FontSubFamilies.BoldItalic), CSF(1.33f, 1.22f, 1.07f));

            _fonts.Add(GetKey(FontMetricsFamilies.BookmanOldStyle, FontSubFamilies.Regular), CSF(1.17f, 1.15f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.BookmanOldStyle, FontSubFamilies.Bold), CSF(1.21f, 1.19f, 1.14f));
            _fonts.Add(GetKey(FontMetricsFamilies.BookmanOldStyle, FontSubFamilies.Italic), CSF(1.12f, 1.1f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.BookmanOldStyle, FontSubFamilies.BoldItalic), CSF(1.23f, 1.20f, 1.15f));

            _fonts.Add(GetKey(FontMetricsFamilies.CalistoMT, FontSubFamilies.Regular), CSF(1.23f, 1.12f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalistoMT, FontSubFamilies.Bold), CSF(1.26f, 1.13f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalistoMT, FontSubFamilies.Italic), CSF(1.04f, 1.02f, 0.99f));
            _fonts.Add(GetKey(FontMetricsFamilies.CalistoMT, FontSubFamilies.BoldItalic), CSF(1.12f, 1.04f, 1.02f));

            _fonts.Add(GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Regular), CSF(1.20f, 1.15f, 1.02f, 1.5f));
            _fonts.Add(GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Bold), CSF(1.20f, 1.15f, 1.01f, 1.5f));
            _fonts.Add(GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.Italic), CSF(1.20f, 1.20f, 1.05f, 1.5f));
            _fonts.Add(GetKey(FontMetricsFamilies.TimesNewRoman, FontSubFamilies.BoldItalic), CSF(1.20f, 1.19f, 1.04f, 1.5f));

            _fonts.Add(GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Regular), CSF(1.17f, 1.12f, 1.04f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Bold), CSF(1.17f, 1.11f, 1.04f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.Italic), CSF(1.17f, 1.11f, 1.04f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.CourierNew, FontSubFamilies.BoldItalic), CSF(1.17f, 1.11f, 1.04f, 1.2f));

            _fonts.Add(GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Regular), CSF(1.17f, 1.08f, 1.02f));
            _fonts.Add(GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Bold), CSF(1.16f, 1.09f, 1.04f));
            _fonts.Add(GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.Italic), CSF(1.18f, 1.13f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.LiberationSerif, FontSubFamilies.BoldItalic), CSF(1.18f, 1.08f, 1.05f));

            _fonts.Add(GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Regular), CSF(1.17f, 1.12f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Bold), CSF(1.33f, 1.26f, 1.17f));
            _fonts.Add(GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.Italic), CSF(1.17f, 1.12f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.Verdana, FontSubFamilies.BoldItalic), CSF(1.3f, 1.3f, 1.18f));

            _fonts.Add(GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Regular), CSF(1.18f, 1.08f, 1.06f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Bold), CSF(1.23f, 1.11f, 1.06f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.Italic), CSF(1.19f, 1.08f, 1.06f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.Cambria, FontSubFamilies.BoldItalic), CSF(1.25f, 1.07f, 1.04f, 1.3f));

            _fonts.Add(GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Regular), CSF(1.15f, 1.10f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Bold), CSF(1.35f, 1.26f, 1.20f));
            _fonts.Add(GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.Italic), CSF(1.13f, 1.13f, 1.1f));
            _fonts.Add(GetKey(FontMetricsFamilies.Georgia, FontSubFamilies.BoldItalic), CSF(1.39f, 1.31f, 1.23f));

            _fonts.Add(GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Regular), CSF(1.22f, 1.10f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Bold), CSF(1.17f, 1.05f, 1.04f));
            _fonts.Add(GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.Italic), CSF(1.20f, 1.13f, 1.11f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.Corbel, FontSubFamilies.BoldItalic), CSF(1.16f, 1.04f, 1.02f));

            _fonts.Add(GetKey(FontMetricsFamilies.Garamond, FontSubFamilies.Regular), CSF(1.42f, 1.13f, 1.05f, 0.8f));
            _fonts.Add(GetKey(FontMetricsFamilies.Garamond, FontSubFamilies.Bold), CSF(1.49f, 1.21f, 1.15f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.Garamond, FontSubFamilies.Italic), CSF(1.18f, 0.83f, 0.98f, 2.5f));
            _fonts.Add(GetKey(FontMetricsFamilies.Garamond, FontSubFamilies.BoldItalic), CSF(1.30f, 1.08f, 1.01f, 0.9f));

            _fonts.Add(GetKey(FontMetricsFamilies.GillSansMT, FontSubFamilies.Regular), CSF(1.25f, 1.11f, 1.07f));
            _fonts.Add(GetKey(FontMetricsFamilies.GillSansMT, FontSubFamilies.Bold), CSF(1.38f, 1.3f, 1.19f, 1.75f));
            _fonts.Add(GetKey(FontMetricsFamilies.GillSansMT, FontSubFamilies.Italic), CSF(1.14f, 1.08f, 1.04f, 1.3f));
            _fonts.Add(GetKey(FontMetricsFamilies.GillSansMT, FontSubFamilies.BoldItalic), CSF(1.28f, 1.23f, 1.12f, 1.1f));

            _fonts.Add(GetKey(FontMetricsFamilies.Impact, FontSubFamilies.Regular), CSF(1.23f, 1.13f, 1.07f, 1.1f));
            _fonts.Add(GetKey(FontMetricsFamilies.Impact, FontSubFamilies.Bold), CSF(1.18f, 1.15f, 1.11f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.Impact, FontSubFamilies.Italic), CSF(1.23f, 1.13f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.Impact, FontSubFamilies.BoldItalic), CSF(1.16f, 1.16f, 1.1f));

            _fonts.Add(GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Regular), CSF(1.16f, 1.13f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Bold), CSF(1.19f, 1.13f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.Italic), CSF(1.21f, 1.15f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturyGothic, FontSubFamilies.BoldItalic), CSF(1.27f, 1.15f, 1.08f));

            _fonts.Add(GetKey(FontMetricsFamilies.CenturySchoolbook, FontSubFamilies.Regular), CSF(1.16f, 1.13f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturySchoolbook, FontSubFamilies.Bold), CSF(1.34f, 1.26f, 1.17f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturySchoolbook, FontSubFamilies.Italic), CSF(1.18f, 1.12f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.CenturySchoolbook, FontSubFamilies.BoldItalic), CSF(1.32f, 1.25f, 1.15f));

            _fonts.Add(GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Regular), CSF(1.31f, 1.09f, 1.08f));
            _fonts.Add(GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Bold), CSF(1.32f, 1.19f, 1.13f));
            _fonts.Add(GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.Italic), CSF(1.24f, 1.08f, 1.05f));
            _fonts.Add(GetKey(FontMetricsFamilies.Rockwell, FontSubFamilies.BoldItalic), CSF(1.30f, 1.19f, 1.09f));

            _fonts.Add(GetKey(FontMetricsFamilies.RockwellCondensed, FontSubFamilies.Regular), CSF(1.42f, 1.13f, 1.05f, 0.8f));
            _fonts.Add(GetKey(FontMetricsFamilies.RockwellCondensed, FontSubFamilies.Bold), CSF(1.52f, 1.36f, 1.30f, 1.5f));
            _fonts.Add(GetKey(FontMetricsFamilies.RockwellCondensed, FontSubFamilies.Italic), CSF(1.42f, 1.13f, 1.05f, 0.8f));
            _fonts.Add(GetKey(FontMetricsFamilies.RockwellCondensed, FontSubFamilies.BoldItalic), CSF(1.50f, 1.36f, 1.30f, 1.2f));

            _fonts.Add(GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Regular), CSF(1.23f, 1.12f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Bold), CSF(1.21f, 1.14f, 1.13f));
            _fonts.Add(GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.Italic), CSF(1.17f, 1.17f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.TrebuchetMS, FontSubFamilies.BoldItalic), CSF(1.20f, 1.16f, 1.13f));

            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Regular), CSF(1.13f, 1.10f, 1.03f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Bold), CSF(1.30f, 1.16f, 1.12f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.Italic), CSF(1.37f, 1.12f, 1.06f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMT, FontSubFamilies.BoldItalic), CSF(1.20f, 1.12f, 1.07f));

            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Regular), CSF(1.13f, 1.11f, 1.09f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Bold), CSF(1.38f, 1.34f, 1.21f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.Italic), CSF(1.10f, 1.09f, 1.10f, 1.2f));
            _fonts.Add(GetKey(FontMetricsFamilies.TwCenMTCondensed, FontSubFamilies.BoldItalic), CSF(1.38f, 1.32f, 1.2f, 1.2f));
        }

        public float GetScaleFactor(uint key, float width)
        {
            if(!_fonts.ContainsKey(key))
            {
                return 1f;
            }
            var factor = _fonts[key];
            return factor.Calculate(width);
        }
    }
}
