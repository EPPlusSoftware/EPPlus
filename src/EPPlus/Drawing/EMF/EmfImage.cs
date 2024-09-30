using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EmfImage
    {
        internal List<EMR_RECORD> records = new List<EMR_RECORD>();

        uint size = 0;

        internal EmfImage() { }

        internal void Read(string emf)
        {
            using (FileStream fileStream = new FileStream(emf, FileMode.Open, FileAccess.Read))
            {
                using (BinaryReader br = new BinaryReader(fileStream))
                {
                    ReadEmfRecords(br);
                }
            }
        }

        internal void Read(byte[] emf)
        {
            using (Stream emfByteStream = new MemoryStream(emf))
            {
                using (BinaryReader br = new BinaryReader(emfByteStream))
                {
                    ReadEmfRecords(br);
                }
            }
        }

        private void ReadEmfRecords(BinaryReader br)
        {
            while (br.BaseStream.Position < br.BaseStream.Length)
            {
                EMR_RECORD record;
                var TypeValue = br.ReadUInt32();
                switch (TypeValue)
                {
                    case 0x00000001:
                        record = new EMR_HEADER(br, TypeValue);
                        break;
                    case 0x0000000E:
                        record = new EMR_EOF(br, TypeValue);
                        break;
                    case 0x00000016:
                        record = new EMR_SETTEXTALIGN(br, TypeValue);
                        break;
                    case 0x0000004D:
                        record = new EMR_STRETCHBLT(br, TypeValue);
                        break;
                    case 0x00000051:
                        record = new EMR_STRETCHDIBITS(br, TypeValue);
                        break;
                    case 0x0000001E:
                        record = new EMR_INTERSECTCLIPRECT(br, TypeValue);
                        break;
                    case 0x00000052:
                        record = new EMR_EXTCREATEFONTINDIRECTW(br, TypeValue);
                        break;
                    case 0x00000054:
                        record = new EMR_EXTTEXTOUTW(br, TypeValue);
                        break;
                    default:
                        record = new EMR_RECORD(br, TypeValue, true);
                        break;
                }
                records.Add(record);
                size += record.Size;
            }
        }

        internal void SetNewTextInDefaultEMFImage(string Text)
        {
            EmfCalculateTextLength ectl = new EmfCalculateTextLength(Text);
            records.RemoveRange(17, 6);
            records.InsertRange(17, ectl.TextRecords);
        }

        internal void ChangeImage(byte[] Image)
        {
            var record = records[16] as EMR_STRETCHBLT;
            record.ChangeImage(Image);
        }

        internal void Save(string FilePath)
        {
            var header = (EMR_HEADER)records[0];
            var last = (EMR_EOF)records[records.Count - 1];
            var preBytes = header.Bytes;

            header.Bytes = 0;

            foreach (var record in records)
            {
                header.Bytes += record.Size;
            }

            using (FileStream fileStream = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
            {
                using (BinaryWriter br = new BinaryWriter(fileStream))
                {
                    foreach (var record in records)
                    {
                        record.WriteBytes(br);
                    }
                }
            }
        }

        internal byte[] GetBytes()
        {
            var header = (EMR_HEADER)records[0];
            var last = (EMR_EOF)records[records.Count - 1];
            var preBytes = header.Bytes;

            header.Bytes = 0;

            foreach (var record in records)
            {
                header.Bytes += record.Size;
            }

            using (MemoryStream ms = new MemoryStream())
            {
                using (BinaryWriter br = new BinaryWriter(ms))
                {
                    foreach (var record in records)
                    {
                        record.WriteBytes(br);
                    }
                    return ms.ToArray();
                }
            }
        }
    }
}
