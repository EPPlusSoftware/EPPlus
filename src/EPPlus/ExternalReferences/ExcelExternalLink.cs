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
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    public abstract class ExcelExternalLink
    {
        internal ExcelWorkbook _wb;
        internal ExcelExternalLink(ExcelWorkbook wb)
        {
            _wb = wb;
            As = new ExcelExternalLinkAsType(this);            
            Part = null;
            WorkbookElement = null;
        }
        internal ExcelExternalLink(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement)
        {
            _wb = wb;
            As = new ExcelExternalLinkAsType(this);
            Part = part;
            WorkbookElement = workbookElement;
        }
        public abstract eExternalLinkType ExternalLinkType
        {
            get;
        }
        internal abstract void Save(StreamWriter sw);
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
        public int Index
        {
            get
            {
                return _wb.ExternalReferences.GetIndex(this)+1;
            }
        }
        internal static bool HasWebProtocol(string uriPath)
        {
            return uriPath.StartsWith("http:") || uriPath.StartsWith("https:") || uriPath.StartsWith("ftp:") || uriPath.StartsWith("ftps:");
        }
        protected internal StringBuilder _errors = new StringBuilder();
        public string ErrorsLog
        {
            get
            {
                return _errors.ToString();
            }
        }

    }
}