using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Utils
{
    internal static class ChecksumCalculator
    {

        /// <summary>
        /// Calculates the checksum for a table according to OpenType spec.
        /// Pads the data to a multiple of 4 bytes and interprets values as big-endian.
        /// </summary>

        public static uint CalculateTableChecksum(byte[] data, string tag)
        {

            if (tag == "head" && data.Length >= 12)
            {
                data = (byte[])data.Clone(); // Make a copy so we don't modify original
                for (int i = 8; i < 12; i++)
                {
                    data[i] = 0;
                }
            }

            uint sum = 0;
            int length = data.Length;
            int paddedLength = ((length + 3) / 4) * 4; // Round up to multiple of 4
            for (int i = 0; i < paddedLength; i += 4)
            {
                uint value = 0;
                for (int b = 0; b < 4; b++)
                {
                    int index = i + b;
                    byte byteValue = (index < length) ? data[index] : (byte)0;
                    value = (value << 8) | byteValue; // Big-endian shift
                }
                sum += value;
            }
            return sum;
        }


        /// <summary>
        /// Calculates the checksum for the entire font file.
        /// The head.checkSumAdjustment field must be set to 0 before calling this.
        /// </summary>
        public static uint CalculateFontChecksum(byte[] fontData)
        {
            uint sum = 0;
            int length = fontData.Length;
            int paddedLength = ((length + 3) / 4) * 4;

            for (int i = 0; i < paddedLength; i += 4)
            {
                uint value = 0;
                for (int b = 0; b < 4; b++)
                {
                    int index = i + b;
                    byte byteValue = (index < length) ? fontData[index] : (byte)0;
                    value = (value << 8) | byteValue;
                }
                sum += value;
            }
            return sum;
        }

    }
}
