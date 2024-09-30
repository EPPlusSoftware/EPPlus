using System;

namespace OfficeOpenXml.Drawing.EMF
{
    [Flags]
    internal enum TextAlignmentModeFlags
    {
        TA_NOUPDATECP = 0x0000,
        TA_LEFT = 0x0000,
        TA_TOP = 0x0000,
        TA_UPDATECP = 0x0001,
        TA_RIGHT = 0x0002,
        TA_CENTER = 0x0006,
        TA_BOTTOM = 0x0008,
        TA_BASELINE = 0x0018,
        TA_RTLREADING = 0x0100,
    }

    [Flags]
    internal enum VerticalTextAlignmentModeFlags
    {
        VTA_TOP = 0x0000,
        VTA_RIGHT = 0x0000,
        VTA_BOTTOM = 0x0002,
        VTA_CENTER = 0x0006,
        VTA_LEFT = 0x0008,
        VTA_BASELINE = 0x0018
    }
}
