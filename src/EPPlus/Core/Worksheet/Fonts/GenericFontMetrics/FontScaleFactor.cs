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
            if (width < (100 * _sizeFactor)) return Adjustment(width, (100 * _sizeFactor), (25 * _sizeFactor), _small);
            else if (width < (200 * _sizeFactor)) return _medium;
            else return _large;
        }

        private float Adjustment(float v, float upper, float lower, float originalFactor)
        {
            var val = v > lower ? (v < upper ? v : upper) : lower;
            var factor = upper - val;
            var factorAdjustment = factor/(upper-lower);
            return originalFactor * (1f + 0.1f * factorAdjustment);
        }
    }
}
