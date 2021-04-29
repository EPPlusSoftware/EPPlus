/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/28/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Core.ExternalReferences
{
    public abstract class ExcelExternalLink
    {
        internal ExcelExternalLink(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement)
        {
            _wb = wb;
            Part = part;
            WorkbookElement = workbookElement;
            
            As = new ExcelExternalLinkAsType(this);
        }
        public abstract eExternalLinkType ExternalLinkType
        {
            get;
        }
        internal abstract void Save(StreamWriter sw);
        internal ExcelWorkbook _wb;
        internal XmlElement WorkbookElement
        {
            get;
            set;
        }

        internal ZipPackagePart Part
        {
            get;
            set;
        }
        public ExcelExternalLinkAsType As
        {
            get;
        }
        public override string ToString()
        {
            return ExternalLinkType.ToString();
        }
    }
}