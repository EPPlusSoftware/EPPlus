/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables;

public class Version16Dot16 : FontTableElement
{
    public Version16Dot16(int value)
    {
        RawValue = value;
        Major = (value >> 16) & 0xFFFF;
        Minor = value & 0xFFFF;
        FloatValue = Major + (Minor / 65536f);
    }

    /// <summary>
    /// Raw 32-bit packed value
    /// </summary>
    public int RawValue { get; }

    /// <summary>
    /// Major version (upper 16 bits)
    /// </summary>
    public int Major { get; }

    /// <summary>
    /// Minor version (lower 16 bits)
    /// </summary>
    public int Minor { get; }

    /// <summary>
    /// Version as float (e.g. 1.5f for 0x00018000)
    /// </summary>
    public float FloatValue { get; }

    public override string ToString()
    {
        return $"{Major}.{FloatValue - Major:F4}";
    }

    internal override void Serialize(FontsBinaryWriter writer)
    {
        writer.WriteInt32BigEndian(RawValue);
    }
}