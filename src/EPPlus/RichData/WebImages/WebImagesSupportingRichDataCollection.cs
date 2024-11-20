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
using OfficeOpenXml.RichData.RichValueArrays;
using OfficeOpenXml.RichData.RichValues.WebImages;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.WebImages
{
    internal class WebImagesSupportingRichDataCollection : IndexedCollection<WebImagesSupportingRichData>
    {
        const string PART_URI_PATH = "/xl/richData/rdRichValueWebImage.xml";
        private readonly Uri _uri;
        private ExcelWorkbook _wb;
        private readonly ExcelRichData _richData;
        private readonly RichDataIndexStore _indexStore;
        ZipPackagePart _part;
        internal ZipPackagePart Part { get { return _part; } }

        public WebImagesSupportingRichDataCollection(ExcelWorkbook wb, ExcelRichData richData) : base(wb.IndexStore, RichDataEntities.WebImage)
        {
            _wb = wb;
            _richData = richData;
            _indexStore = wb.IndexStore;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataWebImage).FirstOrDefault();
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

        private string ExtLstXml
        {
            get;
            set;
        }

        private void LoadPart(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(_uri))
            {
                _part = wb._package.ZipPackage.GetPart(_uri);
                ReadXml(_part.GetStream());
            }
        }

        internal void CreatePart()
        {
            if (_part == null)
            {
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataWebImage);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataWebImage);
            }
            _part.SaveHandler = Save;
        }

        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("webImageSrd"))
                {
                    var array = new WebImagesSupportingRichData(_wb, _part, xr);
                    Add(array);
                }
                else if(xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadElementContentAsString();
                }
                else if(xr.IsEndElementWithName("webImagesSrd"))
                {
                    break;
                }
            }
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<webImagesSrd xmlns=\"{Schemas.schemaWebImage}\" xmlns:r=\"{ExcelPackage.schemaRelationships}\">");
            foreach (var item in this)
            {
                item.WriteXml(sw);
            }
            sw.Write("</webImagesSrd>");
            sw.Flush();
        }
    }
}
