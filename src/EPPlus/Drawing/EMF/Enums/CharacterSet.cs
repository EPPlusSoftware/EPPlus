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
namespace OfficeOpenXml.Drawing.EMF
{
    enum CharacterSet
    {
        ANSI_CHARSET = 0x00000000,
        DEFAULT_CHARSET = 0x00000001,
        SYMBOL_CHARSET = 0x00000002,
        MAC_CHARSET = 0x0000004D,
        SHIFTJIS_CHARSET = 0x00000080,
        HANGUL_CHARSET = 0x00000081,
        JOHAB_CHARSET = 0x00000082,
        GB2312_CHARSET = 0x00000086,
        CHINESEBIG5_CHARSET = 0x00000088,
        GREEK_CHARSET = 0x000000A1,
        TURKISH_CHARSET = 0x000000A2,
        VIETNAMESE_CHARSET = 0x000000A3,
        HEBREW_CHARSET = 0x000000B1,
        ARABIC_CHARSET = 0x000000B2,
        BALTIC_CHARSET = 0x000000BA,
        RUSSIAN_CHARSET = 0x000000CC,
        THAI_CHARSET = 0x000000DE,
        EASTEUROPE_CHARSET = 0x000000EE,
        OEM_CHARSET = 0x000000FF
    }
}
