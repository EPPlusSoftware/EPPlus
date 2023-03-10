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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadata : XmlHelper
    {
        private ExcelWorkbook _workbook;

        public ExcelMetadata(ExcelWorkbook workbook, XmlNamespaceManager nsm) : base(nsm) 
        {
            _workbook = workbook;
            var p = _workbook._package;
            var rel = _workbook.Part.GetRelationshipsByType(ExcelPackage.schemaMetadata).FirstOrDefault();
            if(rel!=null)
            {
                Part = p.ZipPackage.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                ReadMetadata();
            }
            TopNode = MetadataXml.DocumentElement;
        }

        private void ReadMetadata()
        {
            MetadataXml = new XmlDocument();
            LoadXmlSafe(MetadataXml, Part.GetStream());
            LoadTypes();
        }

        internal List<ExcelMetadataType> MetadataTypes { get; } = new List<ExcelMetadataType>();
        internal List<ExcelFutureMetadataType> FutureMetadataTypes{ get; } = new List<ExcelFutureMetadataType>();
        internal List<ExcelCellMetadata> CellMetadata { get; } = new List<ExcelCellMetadata>();
        private void LoadTypes()
        {
            foreach (XmlElement mdNode in GetNodes("d:metadataTypes"))
            {
                MetadataTypes.Add(new ExcelMetadataType(NameSpaceManager, mdNode));
            }

            foreach (XmlElement cellFmdNode in GetNodes("d:futureMetadata[Name='XLDAPR']/d:bk/d:extLst/d:ext[uri='{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}']/xda:dynamicArrayProperties"))
            {
                FutureMetadataTypes.Add(new ExcelFutureMetadataType(NameSpaceManager, cellFmdNode));
            }

            foreach (XmlElement cellMdNode in GetNodes("d:cellMetadata"))
            {
                CellMetadata.Add(new ExcelCellMetadata(NameSpaceManager, cellMdNode));
            }
        }

        internal bool IsFormulaDynamic(int cm)
        {
            if(cm <= CellMetadata.Count)
            {
                var cellMetadata = CellMetadata[cm-1];
                var metadataType = MetadataTypes[cellMetadata.MetadataRecordTypeIndex];
                if(metadataType.Name == "XLDAPR")
                {
                    return FutureMetadataTypes[cellMetadata.MetadataValueTypeIndex].IsDynamicArray;
                }
            }
            return false;
        }

        public XmlDocument MetadataXml { get; private set; }
        public ZipPackagePart Part { get; set; }
    }
}