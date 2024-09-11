﻿using System.Collections.Generic;
using System.IO;

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
                            //case 0x00000016:
                            //    record = new EMR_SETTEXTALIGN(br, TypeValue);
                            //    break;
                            //case 0x0000004D:
                            //    record = new EMR_STRETCHBLT(br, TypeValue);
                            //    break;
                            //case 0x00000051:
                            //    record = new EMR_STRETCHDIBITS(br, TypeValue);
                            //    break;
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
            }
        }

        public void CreateTextRecord(string Text)
        {
            var record = new EMR_EXTTEXTOUTW(Text);
            records[20] = record;
            //records.Add(record);
        }

        public void Save(string FilePath)
        {
            //var eof = new EMR_EOF();
            //records.Add(eof);
            var header = new EMR_HEADER(records);

            records[0] = header;


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
    }
}
