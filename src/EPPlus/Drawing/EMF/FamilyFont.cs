using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{

    enum FamilyFont
    {
        /// <summary>
        /// Implementation dependent
        /// </summary>
        FF_DONTCARE = 0x00,
        /// <summary>
        /// Variable stroke widths proportional and with serifs
        /// </summary>
        FF_ROMAN = 0x01,
        /// <summary>
        /// Variable stroke widths proportional and sans serifs
        /// </summary>
        FF_SWISS = 0x02,
        /// <summary>
        /// Constant stroke width both with and without serifs
        /// </summary>
        FF_MODERN = 0x03,
        /// <summary>
        /// Font made to look like handwriting
        /// </summary>
        FF_SCRIPT = 0x04,
        /// <summary>
        /// Novelty fonts. That are more decorative e.g. Old English
        /// </summary>
        FF_DECORATIVE = 0x05
    };

    enum Pitch
    {
        /// <summary>
        /// Implementation dependent
        /// </summary>
        DEFAULT_PITCH = 0,
        /// <summary>
        /// All chars occupy same width
        /// </summary>
        FIXED_PITCH = 1,
        /// <summary>
        /// Variable pitch proportional to the glyphs
        /// </summary>
        VARIABLE_PITCH = 2
    }
}

