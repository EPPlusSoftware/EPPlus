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
namespace EPPlus.Fonts.OpenType
{
    public class TableRecord
    {
        public Tag Tag { get; set; }

        public uint Checksum { get; set; }

        public uint Offset { get; set; }

        public uint Length { get; set; }

        public byte[] GetTableBytes(OpenTypeFont font)
        {
            switch(Tag.Value.ToLower())
            {
                case "glyf":
                    return font.GlyfTable?.Serialize() ?? new byte[0];
                case "os/2":
                    return font.Os2Table.Serialize();
                case "cmap":
                    return font.CmapTable.Serialize();
                case "head":
                    return font.HeadTable.Serialize();
                case "hhea":
                    return font.HheaTable.Serialize();
                case "hmtx":
                    return font.HmtxTable.Serialize();
                case "kern":
                    return font.KernTable?.Serialize() ?? new byte[0];
                case "loca":
                    return font.LocaTable.Serialize();
                case "maxp":
                    return font.MaxpTable.Serialize();
                case "name":
                    return font.NameTable.Serialize();
                case "post":
                    return font.PostTable.Serialize();
                default:
                    return new byte[0];
            }
        }
    }
}
