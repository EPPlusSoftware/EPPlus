/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal partial class ExcelMetadataType : XmlHelper
    {
        public ExcelMetadataType(XmlNamespaceManager nsm, XmlElement topNode) : base(nsm, topNode)
        {
            Name = GetXmlNodeString("@name");            
            MinSupportedVersion = GetXmlNodeInt("@minSupportedVersion");
            SetXmlNodeFlag("@ghostRow", MetadataFlags.GhostRow, ref _flags);
            SetXmlNodeFlag("@ghostCol", MetadataFlags.GhostCol, ref _flags);
            SetXmlNodeFlag("@edit", MetadataFlags.Edit, ref _flags);
            SetXmlNodeFlag("@delete", MetadataFlags.Delete, ref _flags);
            SetXmlNodeFlag("@copy", MetadataFlags.Copy, ref _flags);
            SetXmlNodeFlag("@pasteAll", MetadataFlags.PasteAll, ref _flags);
            SetXmlNodeFlag("@pasteFormulas", MetadataFlags.PasteFormulas, ref _flags);
            SetXmlNodeFlag("@pasteValues", MetadataFlags.PasteValues, ref _flags);
            SetXmlNodeFlag("@pasteFormats", MetadataFlags.PasteFormats, ref _flags);
            SetXmlNodeFlag("@pasteComments", MetadataFlags.PasteComments, ref _flags);
            SetXmlNodeFlag("@pasteDataValidation", MetadataFlags.PasteDataValidation, ref _flags);
            SetXmlNodeFlag("@pasteBorders", MetadataFlags.PasteBorders, ref _flags);
            SetXmlNodeFlag("@pasteColWidths", MetadataFlags.PasteColWidths, ref _flags);
            SetXmlNodeFlag("@pasteNumberFormats", MetadataFlags.PasteNumberFormats, ref _flags);
            SetXmlNodeFlag("@merge", MetadataFlags.Merge, ref _flags);
            SetXmlNodeFlag("@splitFirst", MetadataFlags.SplitFirst, ref _flags);
            SetXmlNodeFlag("@rowColShift", MetadataFlags.RowColShift, ref _flags);
            SetXmlNodeFlag("@clearAll", MetadataFlags.ClearAll, ref _flags);
            SetXmlNodeFlag("@clearFormats", MetadataFlags.ClearFormats, ref _flags);
            SetXmlNodeFlag("@clearContents", MetadataFlags.ClearContents, ref _flags);
            SetXmlNodeFlag("@clearComments", MetadataFlags.ClearComments, ref _flags);
            SetXmlNodeFlag("@assign", MetadataFlags.Assign, ref _flags);
            SetXmlNodeFlag("@coerce", MetadataFlags.Coerce, ref _flags);
            SetXmlNodeFlag("@cellMeta", MetadataFlags.CellMeta, ref _flags);
        }
        internal void SetXmlNodeFlag(string path, MetadataFlags flag, ref MetadataFlags value) 
        {            
            if(GetXmlNodeBool(path))
            {
                value |= flag;
            }
            else
            {
                value &= ~flag;
            }
        }

        public string Name 
        {
            get;
            private set;
        }
        public int MinSupportedVersion
        {
            get;
            private set;
        }
        MetadataFlags _flags = 0;
        public MetadataFlags Flags
        {
            get
            {
                return _flags;
            }
        }
    }
}