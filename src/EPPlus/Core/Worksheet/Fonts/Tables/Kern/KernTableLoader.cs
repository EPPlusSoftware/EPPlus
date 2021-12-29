/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Kern
{
    public class KernTableLoader : TableLoader<KernTable>
    {
        public KernTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Kern)
        {
        }

        protected override KernTable LoadInternal()
        {
            var v = _reader.ReadUInt16BigEndian();
            var nt = _reader.ReadUInt16BigEndian();
            var table = new KernTable
            {
                version = v,
                nTables = Convert.ToUInt16(nt)
            };
            var subTables = new List<KernSubTable>();
            var nextTablePos = _reader.BaseStream.Position;
            ushort nPairs = 0;
            for (var x = 0; x < table.nTables; x++)
            {
                var subTable = new KernSubTable
                {
                    version = _reader.ReadUInt16BigEndian(),
                    length = _reader.ReadUInt16BigEndian(),
                    coverage = new KernCoverage(_reader)
                };
                nextTablePos += subTable.length;
                if(subTable.coverage.Format == 0)
                {
                    var format0Table = new KernSubTableFormat0(_reader);
                    if (format0Table.nPairs > 0)
                    {
                        var pairs = new KerningPair[format0Table.nPairs];
                        for(var pIx = 0; pIx < format0Table.nPairs; pIx++)
                        {
                            pairs[pIx] = new KerningPair(_reader);
                            nPairs++;
                        }
                        format0Table.Pairs = pairs;
                    }
                    subTable.Format0Subtable = format0Table;
                    subTables.Add(subTable);
                }
                _reader.BaseStream.Position = nextTablePos;
            }
            table.SubTables = subTables.ToArray();
            table.NumberOfFormat0Tables = nPairs;
            return table;
        }
    }
}
