using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{

    [Flags]
    public enum CompositeGlyphFlags : ushort
    {
        ARG_1_AND_2_ARE_WORDS = 0x0001,
        ARGS_ARE_XY_VALUES = 0x0002,
        ROUND_XY_TO_GRID = 0x0004,
        WE_HAVE_A_SCALE = 0x0008,
        MORE_COMPONENTS = 0x0020,
        WE_HAVE_AN_X_AND_Y_SCALE = 0x0040,
        WE_HAVE_A_TWO_BY_TWO = 0x0080,
        WE_HAVE_INSTRUCTIONS = 0x0100,
        USE_MY_METRICS = 0x0200,
        OVERLAP_COMPOUND = 0x0400,
        SCALED_COMPONENT_OFFSET = 0x0800,
        UNSCALED_COMPONENT_OFFSET = 0x1000
    }

}
