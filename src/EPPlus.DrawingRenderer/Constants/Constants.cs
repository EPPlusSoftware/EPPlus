using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer
{
    internal class Constants
    {
        internal const float STANDARD_DPI = 96;
        /// <summary>
        /// The ratio between EMU and Pixels
        /// </summary>
        public const int EMU_PER_PIXEL = 9525;
        /// <summary>
        /// The ratio between EMU and Points
        /// </summary>
        public const int EMU_PER_POINT = 12700;
        /// <summary>
        /// The ratio between EMU and centimeters
        /// </summary>
        public const int EMU_PER_CM = 360000;
        /// <summary>
        /// The ratio between EMU and millimeters
        /// </summary>
        public const int EMU_PER_MM = 3600000;
        /// <summary>
        /// The ratio between EMU and US Inches
        /// </summary>
        public const int EMU_PER_US_INCH = 914400;
        /// <summary>
        /// The ratio between EMU and pica
        /// </summary>
        public const int EMU_PER_PICA = EMU_PER_US_INCH / 6;

    }
}
