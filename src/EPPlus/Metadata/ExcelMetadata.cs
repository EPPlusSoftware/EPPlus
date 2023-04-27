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
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadata
    {
        private ExcelWorkbook _workbook;
        private ZipPackagePart _part;
        private Uri _uri;
        
        //Preserve xml variables
        private string _metadataStringsXml;
        private string _metadataCount;
        private string _mdxMetadataXml;
        private string _mdxMetadataCount;
        public string _extLstXml;
        public ExcelMetadata(ExcelWorkbook workbook)
        {
            _workbook = workbook;
            var p = _workbook._package;
            var rel = _workbook.Part.GetRelationshipsByType(ExcelPackage.schemaMetadata).FirstOrDefault();
            if(rel!=null)
            {
                _uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                _part = p.ZipPackage.GetPart(_uri);
                ReadMetadata(_part.GetStream());
            }
            else
            {
                _uri = new Uri("/xl/metadata.xml", UriKind.Relative);
            }
        }

        private void ReadMetadata(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while(xr.Read())
            {
                if(xr.NodeType== XmlNodeType.Element)
                {
                    switch (xr.Name)
                    {
                        case "metadataTypes":
                            ReadMetadataTypes(xr);
                            break;
                        case "metadataStrings":
                            //Currently not used. Preserve.
                            _metadataStringsXml = xr.ReadInnerXml();
                            _metadataCount = xr.GetAttribute("count");
                            break;
                        case "mdxMetadata":
                            //Currently not used. Preserve.
                            _mdxMetadataXml = xr.ReadInnerXml();
                            _mdxMetadataCount = xr.GetAttribute("count");
                            break;
                        case "futureMetadata":
                            ReadFutureMetadata(xr);
                            break;
                        case "cellMetadata":
                            ReadMetadataItems(xr, xr.Name, CellMetadata);
                            break;
                        case "valueMetadata":
                            ReadMetadataItems(xr, xr.Name, ValueMetadata);
                            break;
                        case "extLst":
                            _extLstXml = xr.ReadInnerXml();
                            break;
                    }

                }
            }
        }

        private void ReadMetadataItems(XmlReader xr, string elementName, List<ExcelMetadataItem> collection)
        {
            xr.Read();
            while(xr.IsEndElementWithName(elementName) ==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while(xr.IsEndElementWithName("bk")==false)
                    {
                        collection.Add(new ExcelMetadataItem(xr));
                    }
                }
                xr.Read();
            }
        }

        private void ReadFutureMetadata(XmlReader xr)
        {
            var item = new ExcelFutureMetadata();
            item.Name = xr.GetAttribute("name"); ;
            xr.Read();
            while (xr.IsEndElementWithName("futureMetadata") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("ext"))
                {
                    switch (xr.GetAttribute("uri"))
                    {
                        case ExtLstUris.DynamicArrayPropertiesUri:
                            FutureMetadataTypes.Add(new ExcelFutureMetadataDynamicArray(xr));
                            break;
                        case ExtLstUris.RichValueDataUri:
                            FutureMetadataTypes.Add(new ExcelFutureMetadataRichData(xr));
                            break;
                    }                    
                }
                xr.Read();
            }
        }
        private void ReadMetadataTypes(XmlReader xr)
        {            
            xr.Read();
            while(xr.IsEndElementWithName("metadataTypes")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("metadataType"))
                {
                    var item = new ExcelMetadataType(xr);
                    MetadataTypes.Add(item);
                }
                xr.Read();
            }
        }

        internal void CreateDefaultXml()
        {
            MetadataTypes.Add(new ExcelMetadataType() { Name = "XLDAPR", MinSupportedVersion=120000, Flags=MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta });
            FutureMetadataTypes.Add(new ExcelFutureMetadataDynamicArray(true));
            
            var item = new ExcelMetadataItem();
            item.Records.Add(new ExcelMetadataRecord(1,0));
            CellMetadata.Add(item);
        }
        internal bool HasMetadata()
        {
            return _part==null;
        }
        internal List<ExcelMetadataType> MetadataTypes { get; } = new List<ExcelMetadataType>();
        internal List<ExcelFutureMetadataType> FutureMetadataTypes{ get; } = new List<ExcelFutureMetadataType>();
        internal List<ExcelMetadataItem> CellMetadata { get; } = new List<ExcelMetadataItem>();
        internal List<ExcelMetadataItem> ValueMetadata { get; } = new List<ExcelMetadataItem>();        
        internal bool IsFormulaDynamic(int cm)
        {
            if(cm <= CellMetadata.Count)
            {
                var cellMetadata = CellMetadata[cm - 1];
                var record = cellMetadata.Records.First();
                var metadataType = MetadataTypes[record.RecordTypeIndex - 1];
                if(metadataType.Name == "XLDAPR")
                {
                    return FutureMetadataTypes[record.ValueTypeIndex].AsDynamicArray.IsDynamicArray;
                }
            }
            return false;
        }

        internal void GetDynamicArrayIndex(out int cm)
        {
            if(HasMetadata())
            {
                GetDynamicArrayIndex(out cm);
            }
            else
            {
                CreateDefaultXml();
                cm = 1;
            }
        }

    }
}