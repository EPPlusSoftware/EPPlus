/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
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
