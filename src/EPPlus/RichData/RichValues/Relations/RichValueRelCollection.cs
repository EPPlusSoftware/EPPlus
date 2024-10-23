/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValues.Relations
{
    internal class RichValueRelCollection : IndexedCollection<RichValueRel>
    {
        const string PART_URI_PATH = "/xl/richData/richValueRel.xml";
        Uri _uri;
        private ExcelWorkbook _wb;
        ZipPackagePart _part;

        public RichValueRelCollection(ExcelWorkbook wb) : base(wb.IndexStore, RichDataEntities.RichValueRel)
        {
            _wb = wb;
            _indexStore = wb.IndexStore;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataRelRelationship).FirstOrDefault();
            if (r == null)
            {
                _uri = new Uri(PART_URI_PATH, UriKind.Relative);
            }
            else
            {
                _uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
            }
            LoadPart(wb);
        }

        private readonly RichDataIndexStore _indexStore;

        private void LoadPart(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(_uri))
            {
                _part = wb._package.ZipPackage.GetPart(_uri);
                ReadXml(_part.GetStream());
            }
        }

        internal ZipPackagePart Part { get { return _part; } }

        private void ReadXml(Stream stream)
        {
            //var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            var xml = string.Empty;
            //using var sr = new StreamReader(stream);
            //var s = sr.ReadToEnd();
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("rel"))
                {
                    Add(ReadItem(xr));
                }
            }
        }

        private RichValueRel ReadItem(XmlReader xr)
        {
            var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            var item = new RichValueRel(_wb, _part);

            var id = xr.GetAttribute("r:id");
            var rel = _part.GetRelationship(id);

            item.RelationId = id;
            item.Type = rel.RelationshipType;
            item.TargetUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

            return item;
        }

        private void EnsurePartExists(out bool partNotLoaded)
        {
            partNotLoaded = false;
            if (_part == null)
            {
                if (_wb._package.ZipPackage.PartExists(_uri))
                {
                    _part = _wb._package.ZipPackage.GetPart(_uri);
                    partNotLoaded = true;
                }
                else
                {
                    _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValueRel);
                    _wb.Part.CreateRelationship(_uri, TargetMode.Internal, "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel");
                    _part.SaveHandler = Save;
                }
            }
        }

        internal RichValueRel AddItem(Uri targetUri, string type, IndexEndpoint relationOwner, out IndexRelation rel)
        {
            EnsurePartExists(out bool partNotLoaded);
            if (partNotLoaded)
            {
                ReadXml(_part.GetStream());
            }

            var relationship = _part.CreateRelationship(targetUri, TargetMode.Internal, type);
            var rvRel = new RichValueRel(_wb, _part)
            {
                RelationId = relationship.Id,
                TargetUri = relationship.TargetUri,
                Type = relationship.RelationshipType
            };
            Add(rvRel);
            rel = _indexStore.CreateAndAddRelation(relationOwner, rvRel, IndexType.ZeroBasedPointer);
            return rvRel;
        }

        //public RichValueRel GetItem(string relId, out int ix)
        //{
        //    ix = -1;
        //    var item = Items.FirstOrDefault(x => x.Id == relId);
        //    if (item != null)
        //    {
        //        ix = Items.IndexOf(item);
        //    }
        //    return item;
        //}

        //public RichValueRel GetItem(int relIx)
        //{
        //    if (relIx < 0 || relIx >= Items.Count)
        //    {
        //        throw new ArgumentOutOfRangeException(nameof(relIx));
        //    }
        //    return Items[relIx];
        //}

        internal void SetNewTarget(uint newId, Uri targetUri)
        {
            EnsurePartExists(out bool p);
            var rel = GetItem(newId);
            var relationship = _part.GetRelationship(rel.RelationId);
            relationship.TargetUri = targetUri;
            relationship.Target = targetUri.OriginalString;
            rel.TargetUri = targetUri;
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<richValueRels xmlns=\"{Schemas.schemaRichValueRel}\" xmlns:r=\"{ExcelPackage.schemaRelationships}\">");
            foreach (var item in this)
            {
                item.WriteXml(sw);
            }
            sw.Write("</richValueRels>");
            sw.Flush();
        }
    }
}
