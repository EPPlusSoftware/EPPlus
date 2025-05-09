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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.Utils.FileUtils;
using OfficeOpenXml.Utils.XML;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadata
    {
        private ExcelWorkbook _wb;
        //private readonly ExcelRichData _richData;
        private ZipPackagePart _part;
        private Uri _uri;
        
        //Preserve xml variables
        private string _metadataStringsXml;
        private string _metadataStringCount;
        private string _mdxMetadataXml;
        private string _mdxMetadataCount;
        public string _extLstXml;

       

        internal FutureMetadataDynamicArray FutureMetadataDynamicArray { get; private set; }
        internal uint? DynamicArrayTypeId { get; private set; }
        internal ZipPackagePart Part { get { return _part; } }

        internal EventHandler<ValueMetadataReadEventArgs> ValueMetadataRead;

        internal MetadataDatabase Db { get; private set; }

        internal void OnValueMetadataRead(uint id, uint oneBasedIndex)
        {
            ValueMetadataRead?.Invoke(this, new ValueMetadataReadEventArgs(id, oneBasedIndex));
        }

        public ExcelMetadata(ExcelWorkbook workbook)
        {
            _wb = workbook;
            Db = new MetadataDatabase(workbook.IndexStore);
            var rel = _wb.Part.GetRelationshipsByType(ExcelPackage.schemaMetadata).FirstOrDefault();
            if(rel!=null)
            {
                _uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                _part = _wb._package.ZipPackage.GetPart(_uri);
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
                            _metadataStringCount = xr.GetAttribute("count");
                            _metadataStringsXml = ReadElementContentAsString(xr);
                            break;
                        case "mdxMetadata":
                            //Currently not used. Preserve.
                            _mdxMetadataCount = xr.GetAttribute("count");
                            ReadMdxMetadataItems(xr);
                            break;
                        case "futureMetadata":
                            ReadFutureMetadata(xr);
                            break;
                        case "cellMetadata":
                            ReadCellMetadataItems(xr);
                            break;
                        case "valueMetadata":
                            ReadValueMetadataItems(xr);
                            break;
                        case "extLst":
                            _extLstXml = xr.ReadInnerXml();
                            break;
                    }

                }
            }
        }


        private string ReadElementContentAsString(XmlReader xr)
        {
            if (xr.NodeType != XmlNodeType.Element)
                throw new InvalidOperationException("Current node is not an element.");

            var elementName = xr.Name;
            var sb = new StringBuilder();
            var readNext = true;
            while(true)
            {
                if (readNext) xr.Read();
                if (xr.NodeType == XmlNodeType.EndElement && xr.Name == elementName)
                    break;

                sb.Append(xr.ReadOuterXml());
                readNext = false;
            }
            return sb.ToString();
        }


        private void ReadCellMetadataItems(XmlReader xr)
        {
            xr.Read();
            while(xr.IsEndElementWithName("cellMetadata") ==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while(xr.IsEndElementWithName("bk")==false)
                    {
                        Db.CellMetadata.Add(new ExcelCellMetadataBlock(xr, Db));
                    }
                }
                xr.Read();
            }
            SetDynamicArrayIdIfExists();
        }

        private void ReadMdxMetadataItems(XmlReader xr)
        {
            xr.Read();
            while (xr.IsEndElementWithName("mdxMetadata") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("mdx"))
                {
                    Db.MdxMetadata.Add(new MdxMetadata(xr, _wb.IndexStore));
                }
            }
        }

        private void SetDynamicArrayIdIfExists()
        {
            if (Db.MetadataTypes.TryGetValue(FutureMetadataBase.DYNAMIC_ARRAY_NAME, out ExcelMetadataType type))
            {
                var rc = type.GetFirstIncomingRelByType<ExcelCellMetadataRecord>();
                if(rc != null)
                {
                    var bk = rc.GetFirstOutgoingRelByType<ExcelCellMetadataBlock>();
                    if(bk != null)
                    {
                        DynamicArrayTypeId = bk.Id;
                    }
                }
            }
        }

        private void ReadValueMetadataItems(XmlReader xr)
        {
            xr.Read();
            while (xr.IsEndElementWithName("valueMetadata") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while (xr.IsEndElementWithName("bk") == false)
                    {
                        Db.ValueMetadata.Add(new ExcelValueMetadataBlock(xr, Db));
                    }
                }
                xr.Read();
            }
        }

        private void ReadFutureMetadata(XmlReader xr)
        {
            FutureMetadataBase fd;
            var name = xr.GetAttribute("name");
            if(name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
            {
                fd = new FutureMetadataDynamicArray(xr, Db);
            }
            else if(name == FutureMetadataBase.RICHDATA_NAME)
            {
                fd = new FutureMetadataRichValue(xr, Db);
            }
            else
            {
                fd = new FutureMetadataPreserve(xr, _wb.IndexStore);
            }
            fd.Index = Db.FutureMetadata.Count;
            Db.FutureMetadata.Add(fd);
        }

        private void ReadMetadataTypes(XmlReader xr)
        {            
            xr.Read();
            while(xr.IsEndElementWithName("metadataTypes")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("metadataType"))
                {
                    var item = new ExcelMetadataType(xr, _wb.IndexStore);
                    Db.MetadataTypes.Add(item);
                }
                xr.Read();
            }
        }

        internal void InitRelations(RichDataDatabase richDataDb)
        {
            var fm = Db.FutureMetadata.FirstOrDefault(x => x.Name == FutureMetadataBase.RICHDATA_NAME);
            var fmrv = fm as FutureMetadataRichValue;
            if(fmrv != null)
            {
                for(var ix = 0; ix < fmrv.Blocks.Count; ix++)
                {
                    var bk = fmrv.Blocks[ix];
                    bk.InitRelations(richDataDb);
                    fm.AddRelationTo(bk, IndexType.ZeroBasedPointer);
                }
            }
            fm = Db.FutureMetadata.FirstOrDefault(x => x.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME);
            var fmda = fm as FutureMetadataDynamicArray;
            if(fmda != null)
            {
                for (var ix = 0; ix < fmda.Blocks.Count; ix++)
                {
                    var bk = fmda.Blocks[ix];
                    bk.InitRelations(richDataDb);
                    fm.AddRelationTo(bk, IndexType.ZeroBasedPointer);
                }
            }
        }

        internal int CreateDefaultXmlDynamicArray()
        {
            var mt = new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.DYNAMIC_ARRAY_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta };
            Db.MetadataTypes.Add(mt);
            var daType = FutureMetadataDynamicArray.GetDefault(Db, out uint bkId);
            Db.FutureMetadata.Add(daType);

            var cmBlock = new ExcelCellMetadataBlock(Db);
            DynamicArrayTypeId = cmBlock.Id;
            daType.AddRelationTo(cmBlock);
            cmBlock.AddRelationTo(daType);
            cmBlock.AddRecord(mt.Id, bkId);
            Db.CellMetadata.Add(cmBlock);
            return Db.CellMetadata.Count;
        }

        internal void CreateRichValueMetadata(ExcelRichData richData, ExcelRichValue richValue, out uint valueMetadataBlockId)
        {
            var rvMetadataType = default(ExcelMetadataType);
            if(Db.MetadataTypes.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out ExcelMetadataType mt))
            {
                rvMetadataType = mt;
            }
            else
            {
                rvMetadataType = new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.RICHDATA_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce };
                Db.MetadataTypes.Add(rvMetadataType);
            }
            var rdTypeId = rvMetadataType.Id;
            var rvFutureMetadata = default(FutureMetadataBase);
            if(Db.FutureMetadata.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out FutureMetadataBase fmb))
            {
                rvFutureMetadata = fmb;
            }
            else
            {
                rvFutureMetadata = new FutureMetadataRichValue(FutureMetadataBase.RICHDATA_NAME, Db);
                Db.MetadataTypes.CreateRelation(rvMetadataType, rvFutureMetadata, IndexType.String);
                Db.FutureMetadata.Add(rvFutureMetadata);
            }
            var block = new FutureMetadataRichValueBlock(_wb.IndexStore);
            rvFutureMetadata.Blocks.Add(block);
            rvFutureMetadata.AddRelationTo(block, IndexType.ZeroBasedPointer);
            var mdItem = new ExcelValueMetadataBlock(Db);
            valueMetadataBlockId = mdItem.Id;
            Db.ValueMetadata.Add(mdItem);
            var rel = block.AddRelationTo(richValue, IndexType.ZeroBasedPointer);
            block.RichDataId = rel.To.Id;
            mdItem.AddRecord(rdTypeId, rel.From.Id);
        }

        internal bool IsFormulaDynamic(uint cmId)
        {
            if (DynamicArrayTypeId.HasValue && cmId == DynamicArrayTypeId.Value) return true;
            var bk = Db.CellMetadata.Get(cmId);
            if (bk != null)
            {
                if(bk.Records != null && bk.Records.Any())
                {
                    var rc = bk.Records[0];
                    if (rc.HasOutgoingRelationTo(RichDataEntities.MetadataType))
                    {
                        var type = rc.GetFirstOutgoingRelByType<ExcelMetadataType>();
                        if (type != null && type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
                        {
                            return true;
                        }
                    }
                }
                
            }
            return false;
        }
        internal void GetDynamicArrayId(out uint cmId)
        {
            if(!DynamicArrayTypeId.HasValue || !Db.CellMetadata.Any(x => x.Id == DynamicArrayTypeId.Value && !x.Deleted))
            {
                Db.CellMetadata.ReIndex();
                CreateDefaultXmlDynamicArray();
            }
            cmId  = DynamicArrayTypeId.Value;
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<metadata xmlns=\"{Schemas.schemaMain}\" xmlns:xlrd=\"{Schemas.schemaRichData}\" xmlns:xda=\"{Schemas.schemaDynamicArray}\">");
            WriteMetadataTypes(sw);
            WriteMetadataStrings(sw);
            WriteMdxMetadata(sw);
            WriteFutureMetadata(sw);
            WriteCellMetadataItems(sw, "cellMetadata", Db.CellMetadata);
            WriteValueMetadataItems(sw, "valueMetadata", Db.ValueMetadata);
            sw.Write("</metadata>");
            sw.Flush();

        }

        internal void CreatePart()
        {
            if (_part == null)
            {
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeMetaData);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaMetadata);
            }
            _part.SaveHandler = Save;
        }

        private void WriteValueMetadataItems(StreamWriter sw, string element, ValueMetadataBlockCollection collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count(x => !x.Deleted && x.Records.Any(x => x.IsValid))}\">");
            foreach (var item in collection)
            {
                if (item.Deleted) continue;
                var records = item.Records.Where(x => !x.Deleted);
                if (records.Any(x => x.IsValid))
                {
                    sw.Write("<bk>");
                    foreach (var r in records)
                    {
                        var mtIx = r.MetadataTypeIndex;
                        var fmbIx = r.FutureMetadataBlockIndex ?? r.MdxValueMetadataIndex;
                        if (mtIx == null || fmbIx == null) continue;
                        sw.Write($"<rc t=\"{mtIx}\" v=\"{fmbIx}\"/>");
                    }
                    sw.Write("</bk>");
                }
            }
            sw.Write($"</{element}>");
        }

        private void WriteCellMetadataItems(StreamWriter sw, string element, CellMetadataBlockCollection collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count(x => !x.Deleted && x.Records.Any())}\">");
            foreach(var item in collection)
            {
                var records = item.Records.Where(x => !x.Deleted);
                if (records.Any())
                {
                    sw.Write("<bk>");
                    foreach (var r in records)
                    {
                        sw.Write($"<rc t=\"{r.MetadataTypeIndex}\" v=\"{r.FutureMetadataBlockIndex}\"/>");
                    }
                    sw.Write("</bk>");
                }
            }
            sw.Write($"</{element}>");
        }
        private void WriteFutureMetadata(StreamWriter sw)
        {
            if (Db.FutureMetadata.Count > 0)
            {
                foreach(var fmd in Db.FutureMetadata)
                {
                    if(fmd.Deleted) continue;
                    fmd.Save(sw);
                }
            }
        }

        private void WriteMetadataTypes(StreamWriter sw)
        {
            if(Db.MetadataTypes.Count > 0)
            {
                sw.Write($"<metadataTypes count=\"{Db.MetadataTypes.Count}\">");
                foreach (var metadataType in Db.MetadataTypes)
                {
                    if (metadataType.Deleted) continue;
                    metadataType.WriteXml(sw);
                }
                sw.Write($"</metadataTypes>");
            }
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
            if (Db.MdxMetadata != null && Db.MdxMetadata.Count > 0)
            {
                sw.Write($"<mdxMetadata count=\"{Db.MdxMetadata.Count}\">");
                foreach (var mdx in Db.MdxMetadata)
                {
                    if (mdx.Deleted) continue;
                    mdx.Write(sw);

                }
                sw.Write("</mdxMetadata>");
            }
        }

        internal bool IsDynamicIdByIndex(int index)
        {
            if (index >= Db.CellMetadata.Count) return false;
            var bk = Db.CellMetadata[index];
            if (bk.Records == null || !bk.Records.Any()) return false;
            var rc = bk.Records[0];
            var type = rc.GetFirstOutgoingRelByType<ExcelMetadataType>();
            if (type == null) return false;
            return type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME;
        }

        internal bool IsDynamicArrayById(uint cmId)
        {
            var cm = Db.CellMetadata.Get(cmId);
            if (cm == null) return false;
            var type = cm.GetFirstOutgoingRelByType<ExcelMetadataType>();
            if(type == null) return false;
            return type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME;
        }

        internal bool IsRichData(uint valueMetadataBlockId, out uint? richValueId)
        {
            richValueId = null;
            if (valueMetadataBlockId == 0) return false;
            var valueMetadata = Db.ValueMetadata.Get(valueMetadataBlockId);
            if (valueMetadata == null || !valueMetadata.Records.Any())
            {
                return false;
            }
            var metadataType = valueMetadata.Records[0].GetFirstOutgoingRelByType<ExcelMetadataType>();
            if (metadataType == null || metadataType.Name != FutureMetadataBase.RICHDATA_NAME) return false;
            var futureMetadata = metadataType.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if (futureMetadata == null) return false;
            var fmBlock = valueMetadata.Records[0].GetFirstOutgoingRelByType<FutureMetadataBlock>();
            if (fmBlock != null)
            {
                var rv = fmBlock.GetFirstOutgoingRelByType<ExcelRichValue>();
                if (rv != null)
                {
                    richValueId = rv.Id;
                    return true;
                }
            }
            return false;
        }

    }
}