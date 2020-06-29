/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/29/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelSlicerCache : XmlHelper
    {
        ExcelWorkbook _workbook;
        internal ExcelSlicerCache(ExcelWorkbook workbook, XmlNamespaceManager nameSpaceManager, ZipPackageRelationship r) : base(nameSpaceManager)
        {
            _workbook = workbook;
            CacheRel = r;
            var Part = workbook.Part.Package.GetPart(UriHelper.ResolvePartUri(workbook.WorkbookUri, r.TargetUri));
            SlicerCacheXml = new XmlDocument();
            LoadXmlSafe(SlicerCacheXml, Part.GetStream());
            TopNode = SlicerCacheXml.DocumentElement;
        }
        internal ZipPackageRelationship CacheRel{ get; }
        internal ZipPackagePart Part { get; }
        public XmlDocument SlicerCacheXml { get; }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
        }
        public string SourceName
        {
            get
            {
                return GetXmlNodeString("@sourceName");
            }
        }
        public eSlicerSourceType SourceType
        {
            get;
        }
    }
}
