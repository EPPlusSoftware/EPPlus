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