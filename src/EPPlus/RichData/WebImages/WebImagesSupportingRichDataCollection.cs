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
using OfficeOpenXml.RichData.RichValues.Relations;
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
        private Dictionary<string, uint> _cachedImages = new Dictionary<string, uint>();
        private Dictionary<string, string> _blipRelations = new Dictionary<string, string>();
        private Dictionary<string, string> _addressRelations = new Dictionary<string, string>();
        private Dictionary<string, string> _moreImagesRelations = new Dictionary<string, string>();
        ZipPackagePart _part;
        internal ZipPackagePart Part { get { return _part; } }

        public WebImagesSupportingRichDataCollection(ExcelWorkbook wb) : base(wb.IndexStore, RichDataEntities.WebImage)
        {
            _wb = wb;
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

        private string GetKey(Uri blipUri, Uri moreImagesUri, Uri addressUri)
        {
            var sb = new StringBuilder();
            if(blipUri != null && !string.IsNullOrEmpty(blipUri.OriginalString))
            {
                sb.AppendFormat("blip:{0}", blipUri.OriginalString);
            }
            if(moreImagesUri != null && !string.IsNullOrEmpty(moreImagesUri.OriginalString))
            {
                sb.AppendFormat("-mi:{0}", moreImagesUri.OriginalString);
            }
            if(addressUri != null && !string.IsNullOrEmpty(addressUri.OriginalString))
            {
                sb.AppendFormat("-a:{0}", addressUri.OriginalString);
            }
            return sb.ToString();
        }

        public override void Add(WebImagesSupportingRichData item)
        {
            var key = GetKey(item.Blip, item.MoreImagesAddress, item.Address);
            if(!_cachedImages.ContainsKey(key))
            {
                base.Add(item);
                _cachedImages[key] = item.Id;
            }
        }

        public bool TryGet(Uri blipUri, Uri moreImagesUri, Uri addressUri, out uint id)
        {
            var key = GetKey(blipUri, moreImagesUri, addressUri);
            if(_cachedImages.ContainsKey(key))
            {
                id = _cachedImages[key];
                return true;
            }
            id = 0;
            return false;
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
                    _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataWebImage);
                    _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataWebImage);
                    _part.SaveHandler = Save;
                }
            }
        }

        internal WebImagesSupportingRichData AddItem(Uri blipUri, Uri addressUri, Uri moreImagesUri, IndexEndpoint relationOwner, out IndexRelation rel)
        {
            EnsurePartExists(out bool partNotLoaded);
            if (partNotLoaded)
            {
                ReadXml(_part.GetStream());
            }

            string blipRelId;
            if(_blipRelations.ContainsKey(blipUri.OriginalString))
            {
                blipRelId = _blipRelations[blipUri.OriginalString];
            }
            else
            {
                var blipRel = _part.CreateRelationship(blipUri, TargetMode.Internal, ExcelPackage.schemaImage);
                blipRelId = blipRel.Id;
                _blipRelations[blipUri.OriginalString] = blipRelId;
            }
            string addressRelId;
            if (_addressRelations.ContainsKey(addressUri.OriginalString))
            {
                addressRelId = _addressRelations[addressUri.OriginalString];
            }
            else
            {
                var addressRel = _part.CreateRelationship(addressUri, TargetMode.External, ExcelPackage.schemaHyperlink);
                addressRelId = addressRel.Id;
                _addressRelations[addressUri.OriginalString] = addressRelId;
            }
            string moreImagesRelId = null;
            if(moreImagesUri != null)
            {
                if (_moreImagesRelations.ContainsKey(moreImagesUri.OriginalString))
                {
                    moreImagesRelId = _moreImagesRelations[moreImagesUri.OriginalString];
                }
                else
                {
                    var moreImagesRel = _part.CreateRelationship(moreImagesUri, TargetMode.External, ExcelPackage.schemaHyperlink);
                    moreImagesRelId = moreImagesRel.Id;
                    _moreImagesRelations[moreImagesUri.OriginalString] = moreImagesRelId;
                }
            }
           
            var image = new WebImagesSupportingRichData(_wb, _part)
            {
                BlipRelationId = blipRelId,
                AddressRelationId = addressRelId,
                MoreImagesRelationId = moreImagesRelId,
            };
            Add(image);

            rel = _wb.IndexStore.CreateAndAddRelation(relationOwner, image, IndexType.ZeroBasedPointer);
            return image;
        }

        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("webImageSrd"))
                {
                    var webImage = new WebImagesSupportingRichData(_wb, _part, xr);
                    Add(webImage);
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
