using System;
using System.IO;

namespace FontLab1
{
    internal class Tag
    {
        public Tag(BinaryReader reader)
        {
            var b1 = reader.ReadByte();
            var b2 = reader.ReadByte();
            var b3 = reader.ReadByte();
            var b4 = reader.ReadByte();

            var c1 = Convert.ToChar(b1);
            var c2 = Convert.ToChar(b2);
            var c3 = Convert.ToChar(b3);
            var c4 = Convert.ToChar(b4);

            Value = new string(new char[] { c1, c2, c3, c4 });
        }

        public string Value { get; private set; }
    }
}
