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
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        
        //Preserve xml variables
        private string _metadataStringsXml;
        private string _metadataStringCount;
        private string _mdxMetadataXml;
        private string _mdxMetadataCount;
        public string _extLstXml;
        public ExcelMetadata(ExcelWorkbook workbook)
        {
            _wb = workbook;
            var p = _wb._package;
            var rel = _wb.Part.GetRelationshipsByType(ExcelPackage.schemaMetadata).FirstOrDefault();
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
                            _metadataStringCount = xr.GetAttribute("count");
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
            item.Name = xr.GetAttribute("name");
            FutureMetadata.Add(item);
            xr.Read();
            while (xr.IsEndElementWithName("futureMetadata") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("ext"))
                {
                    switch (xr.GetAttribute("uri"))
                    {
                        case ExtLstUris.DynamicArrayPropertiesUri:
                            item.Types.Add(new ExcelFutureMetadataDynamicArray(xr));
                            break;
                        case ExtLstUris.RichValueDataUri:
                            item.Types.Add(new ExcelFutureMetadataRichData(xr));
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

        internal int CreateDefaultXml()
        {
            MetadataTypes.Add(new ExcelMetadataType() { Name = "XLDAPR", MinSupportedVersion=120000, Flags=MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta });
            FutureMetadata.Add(new ExcelFutureMetadata() { Name="XLDAPR" });
            FutureMetadata[0].Types.Add(new ExcelFutureMetadataDynamicArray(true));
            
            var item = new ExcelMetadataItem();
            item.Records.Add(new ExcelMetadataRecord(1,0));
            CellMetadata.Add(item);
            return CellMetadata.Count;
        }
        internal bool HasMetadata()
        {
            return MetadataTypes.Count==0;
        }
        internal List<ExcelMetadataType> MetadataTypes { get; } = new List<ExcelMetadataType>();
        internal List<ExcelFutureMetadata> FutureMetadata{ get; } = new List<ExcelFutureMetadata>();
        internal List<ExcelMetadataItem> CellMetadata { get; } = new List<ExcelMetadataItem>();
        internal List<ExcelMetadataItem> ValueMetadata { get; } = new List<ExcelMetadataItem>();        
        internal bool IsFormulaDynamic(int cm)
        {
            if(cm <= CellMetadata.Count)
            {
                var cellMetadata = CellMetadata[cm - 1];
                var record = cellMetadata.Records.First();
                var metadataType = MetadataTypes[record.RecordTypeIndex - 1];
                if (metadataType.Name == "XLDAPR")
                {
                    return FutureMetadata.Find(x => x.Name == "XLDAPR").Types[record.ValueTypeIndex].AsDynamicArray.IsDynamicArray;
                }
            }
            return false;
        }
        internal bool IsSpillError(int vm)
        {
            return GetErrorType(vm) == 8;
        }
        internal bool IsCalcError(int vm)
        {
            return GetErrorType(vm) == 13;
        }

        internal int GetErrorType(int vm)
        {
            var valueMetadata = ValueMetadata[vm - 1];
            var record = valueMetadata.Records.First();
            var metadataType = MetadataTypes[record.RecordTypeIndex - 1];
            if (metadataType.Name == "XLRICHVALUE")
            {
                var ix = FutureMetadata.Find(x => x.Name == "XLDAPR").Types[record.ValueTypeIndex].AsRichData.Index;
                var rd = _wb.RichData.Values.Items[ix];
                var fieldIx = rd.Structure.Keys.FindIndex(x => x.Name == "errorType");
                if (fieldIx >= 0)
                {
                    return int.Parse(rd.Values[fieldIx]);
                }
            }
            return -1;
        }
        internal void GetDynamicArrayIndex(out int cm)
        {
            if(HasMetadata())
            {
                cm=CreateDefaultXml();                
            }
            else
            {
                var tIx = FutureMetadata.FindIndex(x => x.Name == "XLDAPR")+1;
                if (tIx > 0)
                {
                    cm = CellMetadata.FindIndex(x => x.Records.Exists(y => y.RecordTypeIndex == tIx)) + 1;
                    if(cm<=0)
                    {
                        var mtIx = MetadataTypes.FindIndex(x => x.Name == "XLDAPR") + 1;
                        var item = new ExcelMetadataItem();
                        item.Records.Add(new ExcelMetadataRecord(mtIx, tIx));
                        CellMetadata.Add(item);
                        cm = CellMetadata.Count;
                    }
                }
                else
                {
                    cm=CreateDefaultXml();
                }
            }
        }

        internal void Save()
        {
            if (_part == null)
            {
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeMetaData);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaMetadata);
            }

            var stream = _part.GetStream(FileMode.Create);
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<metadata xmlns=\"{Schemas.schemaMain}\" xmlns:xlrd=\"{Schemas.schemaRichData}\" xmlns:xda=\"{Schemas.schemaDynamicArray}\">");
            WriteMetadataTypes(sw);
            WriteMetadataStrings(sw);
            WriteMdxMetadata(sw);
            WriteFutureMetadata(sw);
            WriteMetadataItems(sw, "cellMetadata", CellMetadata);
            WriteMetadataItems(sw, "valueMetadata", ValueMetadata);
            sw.Write("</metadata>");
            sw.Flush();

        }
        private void WriteMetadataItems(StreamWriter sw, string element, List<ExcelMetadataItem> collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count}\">");
            foreach(var item in collection)
            {
                sw.Write("<bk>");
                foreach(var r in item.Records)
                {
                    sw.Write($"<rc t=\"{r.RecordTypeIndex}\" v=\"{r.ValueTypeIndex}\"/>");
                }
                sw.Write("</bk>");
            }
            sw.Write($"</{element}>");
        }
        private void WriteFutureMetadata(StreamWriter sw)
        {
            if (FutureMetadata.Count > 0)
            {
                foreach (var fmd in FutureMetadata)
                {
                    sw.Write($"<futureMetadata name=\"{fmd.Name}\" count=\"1\">");
                    foreach(var t in fmd.Types)
                    {
                        sw.Write($"<bk><extLst><ext uri=\"{t.Uri}\">");
                        t.WriteXml(sw);
                        sw.Write($"</ext></extLst></bk>");
                    }
                    sw.Write($"</futureMetadata>");
                }
            }
        }

        private void WriteMetadataTypes(StreamWriter sw)
        {
            sw.Write($"<metadataTypes count=\"{MetadataTypes.Count}\">");
            foreach(var metadataType in MetadataTypes )
            {
                metadataType.WriteXml(sw);
            }
            sw.Write($"</metadataTypes>");
        }
        private void WriteMetadataStrings(StreamWriter sw)
        {
            if(!string.IsNullOrEmpty(_metadataStringsXml))
            {
                sw.Write($"<metadataStrings count=\"{_metadataStringCount}\">{_metadataStringsXml}</metadataStrings>");
            }
        }
        private void WriteMdxMetadata(StreamWriter sw)
        {
            if (!string.IsNullOrEmpty(_mdxMetadataXml))
            {
                sw.Write($"<mdxMetadata count=\"{_mdxMetadataCount}\">{_mdxMetadataXml}</metadataStrings>");
            }
        }
    }
}