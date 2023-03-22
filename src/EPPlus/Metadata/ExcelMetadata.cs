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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
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
        }

        private void ReadMetadata()
        {
            MetadataXml = new XmlDocument();
            LoadXmlSafe(MetadataXml, Part.GetStream());
            TopNode = MetadataXml.DocumentElement;
            LoadTypes();
        }
        internal void CreateDefaultXml()
        {
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";
            MetadataXml = new XmlDocument();
            LoadXmlSafe(MetadataXml, xml, Encoding.UTF8);
            TopNode = MetadataXml.DocumentElement;
            LoadTypes();

            var metadataUri = new Uri("/xl/metadata.xml", UriKind.Relative);
            Part = _workbook._package.ZipPackage.CreatePart(metadataUri, ContentTypes.contentTypeMetaData);
            _workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_workbook.WorkbookUri, metadataUri), TargetMode.Internal, ExcelPackage.schemaMetadata);
            var stream = Part.GetStream(System.IO.FileMode.Create);
            MetadataXml.Save(stream);
        }
        internal bool HasMetadata()
        {
            return MetadataXml == null;
        }
        internal List<ExcelMetadataType> MetadataTypes { get; } = new List<ExcelMetadataType>();
        internal List<ExcelFutureMetadataType> FutureMetadataTypes{ get; } = new List<ExcelFutureMetadataType>();
        internal List<ExcelCellMetadata> CellMetadata { get; } = new List<ExcelCellMetadata>();
        private void LoadTypes()
        {
            const string MetadataTypesPath = "d:metadataTypes/d:metadataType";
            if (ExistsNode(MetadataTypesPath))
            {
                foreach (XmlElement mdNode in GetNodes(MetadataTypesPath))
                {
                    MetadataTypes.Add(new ExcelMetadataType(NameSpaceManager, mdNode));
                }
            }

            const string FutureMetadataTypesPath = "d:futureMetadata[@name='XLDAPR']/d:bk/d:extLst/d:ext[@uri='{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}']/xda:dynamicArrayProperties";
            if (ExistsNode(FutureMetadataTypesPath))
            {
                foreach (XmlElement cellFmdNode in GetNodes(FutureMetadataTypesPath))
                {
                    FutureMetadataTypes.Add(new ExcelFutureMetadataType(NameSpaceManager, cellFmdNode));
                } 
            }
            const string CellMetaDataPath = "d:cellMetadata/d:bk/d:rc";
            foreach (XmlElement cellMdNode in GetNodes(CellMetaDataPath))
            {
                CellMetadata.Add(new ExcelCellMetadata(NameSpaceManager, cellMdNode));
            }
        }
        internal bool IsFormulaDynamic(int cm)
        {
            if(cm <= CellMetadata.Count)
            {
                var cellMetadata = CellMetadata[cm - 1];
                var metadataType = MetadataTypes[cellMetadata.MetadataRecordTypeIndex - 1];
                if(metadataType.Name == "XLDAPR")
                {
                    return FutureMetadataTypes[cellMetadata.MetadataValueTypeIndex].IsDynamicArray;
                }
            }
            return false;
        }

        internal void GetDynamicArrayIndex(out int cm)
        {
            if(MetadataXml==null)
            {
                CreateDefaultXml();
                cm=1;
                return;
            }
            else
            {
                //TODO Find dynamic Array settings;
                cm = 1;
            }
        }

        public XmlDocument MetadataXml { get; private set; }
        public ZipPackagePart Part { get; set; }
    }
}