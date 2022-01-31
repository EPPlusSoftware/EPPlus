using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal class FontScaleFactor
    {
        public FontScaleFactor(float small, float medium, float large)
            : this(small, medium, large, 1f)
        {

        }

        public FontScaleFactor(float small, float medium, float large, float sizeFactor)
        {
            _small = small;
            _medium = medium;
            _large = large;
            _sizeFactor = sizeFactor;
        }

        private readonly float _small;
        private readonly float _medium;
        private readonly float _large;
        private readonly float _sizeFactor;

        internal float Calculate(float width)
        {
            if (width < (100 * _sizeFactor)) return Adjustment(width, (25 * _sizeFactor), (100 * _sizeFactor), _small, _medium);
            else if (width < (200 * _sizeFactor)) return Adjustment(width, (100 * _sizeFactor), (200 * _sizeFactor), _medium, _large);
            else return _large;
        }

        private float Adjustment(float v, float lowerWidth, float upperWidth, float originalFactorLower, float originalFactorUpper)
        {
            if (v < lowerWidth) return originalFactorLower;
            if (v > upperWidth) return originalFactorLower;
            var f = originalFactorUpper - originalFactorLower;
            var f2 = v / upperWidth;
            return originalFactorLower + (f * f2);
        }
    }
}
