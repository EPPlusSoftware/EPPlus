/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal partial class ExcelMetadataType
    {
        public ExcelMetadataType()
        {

        }
        public ExcelMetadataType(XmlReader xr) 
        {
            Name = xr.GetAttribute("name");            
            MinSupportedVersion = int.Parse(xr.GetAttribute("minSupportedVersion"));
            SetXmlNodeFlag(xr.GetAttribute("ghostRow"), MetadataFlags.GhostRow, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("ghostCol"), MetadataFlags.GhostCol, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("edit"), MetadataFlags.Edit, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("delete"), MetadataFlags.Delete, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("copy"), MetadataFlags.Copy, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteAll"), MetadataFlags.PasteAll, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteFormulas"), MetadataFlags.PasteFormulas, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteValues"), MetadataFlags.PasteValues, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteFormats"), MetadataFlags.PasteFormats, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteComments"), MetadataFlags.PasteComments, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteDataValidation"), MetadataFlags.PasteDataValidation, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteBorders"), MetadataFlags.PasteBorders, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteColWidths"), MetadataFlags.PasteColWidths, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("pasteNumberFormats"), MetadataFlags.PasteNumberFormats, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("merge"), MetadataFlags.Merge, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("splitFirst"), MetadataFlags.SplitFirst, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("rowColShift"), MetadataFlags.RowColShift, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("clearAll"), MetadataFlags.ClearAll, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("clearFormats"), MetadataFlags.ClearFormats, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("clearContents"), MetadataFlags.ClearContents, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("clearComments"), MetadataFlags.ClearComments, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("assign"), MetadataFlags.Assign, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("coerce"), MetadataFlags.Coerce, ref _flags);
            SetXmlNodeFlag(xr.GetAttribute("cellMeta"), MetadataFlags.CellMeta, ref _flags);
        }
        internal void SetXmlNodeFlag(string s, MetadataFlags flag, ref MetadataFlags value) 
        {            
            if(s!=null && (s=="1" || s.Equals("true",StringComparison.OrdinalIgnoreCase)))
            {
                value |= flag;
            }
            else
            {
                value &= ~flag;
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<metadataType name=\"{Name}\" minSupportedVersion=\"{MinSupportedVersion}\" {GetFlagAttributes()} />");
        }

        private string GetFlagAttributes()
        {
            var sb =new StringBuilder();
            foreach(MetadataFlags f in Enum.GetValues(typeof(MetadataFlags)))
            {
                if((f & Flags)==f)
                {
                    sb.Append($" {f.ToEnumString()}=\"1\"");
                }
            }
            return sb.ToString();
        }

        public string Name 
        {
            get;
            set;
        }
        public int MinSupportedVersion
        {
            get;
            set;
        }
        MetadataFlags _flags = 0;
        public MetadataFlags Flags
        {
            get
            {
                return _flags;
            }
            set
            {
                _flags = value;
            }
        }
    }
}