using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMF
    {
        internal List<EMR_RECORD> records = new List<EMR_RECORD>();

        uint size = 0;

        public EMF() { }

        public void Read(string emf)
        {
            using (FileStream fileStream = new FileStream(emf, FileMode.Open, FileAccess.Read))
            {
                using (BinaryReader br = new BinaryReader(fileStream))
                {
                    ReadEmfRecords(br);
                }
            }
        }

        public void Read(byte[] emf)
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

        public void CreateTextRecord(string Text)
        {
            var record = new EMR_EXTTEXTOUTW(Text);
            records[20] = record;
            //records.Add(record);
        }

        public void CreateTextRecord(string Text, int x, int y)
        {
            var record = new EMR_EXTTEXTOUTW(Text, x, y);
            records[20] = record;
            //records.Add(record);
        }

        public void UpdateTextRecord(string Text)
        {
            var textRecord = records[20] as EMR_EXTTEXTOUTW;
            textRecord.Text = Text;
            //records.Add(record);
        }

        public void SetNewText(string Text)
        {
            //remove current text record block
            //create emfcalculatetextlength
            //insert records from emfcalculatetextlength
        }

        public void ChangeTextAlignment(TextAlignmentModeFlags Flags)
        {
            var record = records[8] as EMR_SETTEXTALIGN;
            record.TextAlignmentMode = Flags;
        }

        public void ChangeImage(byte[] Image)
        {
            var record = records[16] as EMR_STRETCHBLT;
            record.ChangeImage(Image);
        }

        public void Save(string FilePath)
        {
            //var eof = new EMR_EOF();
            //records.Add(eof);
            //var header = new EMR_HEADER(records);

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

        public byte[] GetBytes()
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
