using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables
{
    public class Fixed16Dot16 : FontTableElement
    {
        public Fixed16Dot16(int rawValue)
        {
            RawValue = rawValue;
            FloatValue = rawValue / 65536f;
        }

        public Fixed16Dot16(float floatValue)
        {
            FloatValue = floatValue;
            RawValue = (int)(floatValue * 65536f);
        }

        /// <summary>
        /// The raw value in 16.16 fixed-point-format (signed 32-bit integer)
        /// </summary>
        public int RawValue { get; }

        /// <summary>
        /// The value as a float (ex. -12.5, 0.0, 1.25)
        /// </summary>
        public float FloatValue { get; }

        public override string ToString()
        {
            return FloatValue.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteInt32BigEndian(RawValue);
        }

        internal static Fixed16Dot16 ReadFrom(FontsBinaryReader reader)
        {
            int raw = reader.ReadInt32BigEndian();
            return new Fixed16Dot16(raw);
        }
    }
}
