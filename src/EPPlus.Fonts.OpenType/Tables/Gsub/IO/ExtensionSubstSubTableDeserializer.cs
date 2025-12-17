using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class ExtensionSubstSubTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public ExtensionSubstSubTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public FontTableElement Deserialize(long absoluteStart)
        {
            _reader.BaseStream.Seek(absoluteStart, SeekOrigin.Begin);

            ushort format = _reader.ReadUInt16BigEndian(); // Skall vara 1
            ushort lookuptype = _reader.ReadUInt16BigEndian(); // Den riktiga typen (t.ex. 4)
            uint offset = _reader.ReadUInt32BigEndian(); // 32-bitars offset till subtabellen

            long innerSubTableAbsoluteStart = absoluteStart + offset;

            // Här mappar vi om till den loader som faktiskt behövs
            switch (lookuptype)
            {
                case 1:
                    return new SingleSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                case 4:
                    return new LigatureSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                default:
                    return null; // Eller hantera fler typer vid behov
            }
        }
    }
}
