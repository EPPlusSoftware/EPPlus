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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Constants;

namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// Represent an OOXML Zip package.
    /// </summary>
    internal partial class ZipPackage : ZipPackagePartBase, IDisposable
    {
        internal class ContentType
        {
            internal string Name;
            internal bool IsExtension;
            internal string Match;
            public ContentType(string name, bool isExtension, string match)
            {
                Name = name;
                IsExtension = isExtension;
                Match = match;
            }
        }
        Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);
        internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);
        internal char _dirSeparator='0';
        internal ZipPackage()
        {
            AddNew();
        }

        private void AddNew()
        {
            _contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
            _contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
        }
        internal ZipInputStream _zip;
        internal ZipPackage(Stream stream)
        {
            bool hasContentTypeXml = false;
            if (stream == null || stream.Length == 0)
            {
                AddNew();
            }
            else
            {
                var rels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                stream.Seek(0, SeekOrigin.Begin);
                _zip = new ZipInputStream(stream);
                var e = _zip.GetNextEntry();
                if (e == null)
                {
                    throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                }

                while (e != null)
                {
                    GetDirSeparator(e);
                        
                    if (e.UncompressedSize > 0)
                    {
                        if (e.FileName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
                        {
                            var bytes = GetZipEntryAsByteArray(_zip, e);
                            var xml = bytes.GetEncodedString(out Encoding enc);
                            AddContentTypes(xml, enc);
                            hasContentTypeXml = true;
                        }

                        else if (e.FileName.Equals($"_rels{_dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
                        {
                            var bytes = GetZipEntryAsByteArray(_zip, e);
                            var xml = bytes.GetEncodedString(out Encoding enc);
                            ReadRelation(xml, "", enc);
                        }
                        else if (e.FileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                        {
                            var ba = GetZipEntryAsByteArray(_zip, e);
                            var xml = ba.GetEncodedString(out Encoding enc);
                            rels.Add(GetUriKey(e.FileName), xml);
                        }
                        else
                        {
                            ExtractEntryToPart(_zip, e);
                        }                            
                    }
                    e = _zip.GetNextEntry();
                }
                if (_dirSeparator == '0') _dirSeparator = '/';
                foreach (var p in Parts)
                {
                    string name = Path.GetFileName(p.Key);
                    string extension = Path.GetExtension(p.Key);
                    string relFile = string.Format("{0}_rels/{1}.rels", p.Key.Substring(0, p.Key.Length - name.Length), name);
                    if (rels.ContainsKey(relFile))
                    {
                        p.Value.ReadRelation(rels[relFile], p.Value.Uri.OriginalString, Encoding.UTF8);
                    }
                    if (_contentTypes.ContainsKey(p.Key))
                    {
                        p.Value.ContentType = _contentTypes[p.Key].Name;
                    }
                    else if (extension.Length > 1 && _contentTypes.ContainsKey(extension.Substring(1)))
                    {
                        p.Value.ContentType = _contentTypes[extension.Substring(1)].Name;
                    }
                }
                if (!hasContentTypeXml)
                {
                    throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                }
                if (!hasContentTypeXml)
                {
                    throw (new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor."));
                }
            }
        }

        private static byte[] GetZipEntryAsByteArray(ZipInputStream zip, ZipEntry e)
        {
            var b = new byte[e.UncompressedSize];
            var size = zip.Read(b, 0, (int)e.UncompressedSize);
            return b;
        }

        private void ExtractEntryToPart(ZipInputStream zip, ZipEntry e)
        {

            var part = new ZipPackagePart(this, e);

            var rest = e.UncompressedSize;
            if (rest > int.MaxValue)
            {
                part.Stream = null; //Over 2GB, we use the zip stream directly instead.
            }
            else
            {
                const int BATCH_SIZE = 0x100000;
                part.Stream = RecyclableMemory.GetStream();
                while (rest > 0)
                {
                    var bufferSize = rest > BATCH_SIZE ? BATCH_SIZE : (int)rest;
                    var b = new byte[bufferSize];
                    var size = zip.Read(b, 0, bufferSize);
                    part.Stream.Write(b, 0, size);
                    rest -= size;
                }
            }
            Parts.Add(GetUriKey(e.FileName), part);
        }

        private void GetDirSeparator(ZipEntry e)
        {
            if (_dirSeparator == '0')
            {
                if (e.FileName.Contains("\\"))
                {
                    _dirSeparator = '\\';
                }
                else if (e.FileName.Contains("/"))
                {
                    _dirSeparator = '/';
                }
            }
        }

        private void AddContentTypes(string xml, Encoding enc)
        {
            var doc = new XmlDocument();
            XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

            foreach (XmlElement c in doc.DocumentElement.ChildNodes)
            {
                ContentType ct;
                if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
                    _contentTypes.Add(GetUriKey(ct.Match), ct);
                }
                else
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
                    _contentTypes.Add(ct.Match, ct);
                }
            }
        }

#region Methods
        internal ZipPackagePart CreatePart(Uri partUri, string contentType)
        {
            return CreatePart(partUri, contentType, CompressionLevel.Default);
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel, string extension=null)
        {
            if (PartExists(partUri))
            {
                throw (new InvalidOperationException("Part already exist"));
            }

            var part = new ZipPackagePart(this, partUri, contentType, compressionLevel);
            if(string.IsNullOrEmpty(extension))
            {
                _contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
            }
            else
            {
                if (!_contentTypes.ContainsKey(extension))
                {
                    _contentTypes.Add(extension, new ContentType(contentType, true, extension));
                }
            }
            Parts.Add(GetUriKey(part.Uri.OriginalString), part);
            return part;
        }
        internal ZipPackagePart CreatePart(Uri partUri, ZipPackagePart sourcePart)
        {
            var destPart = CreatePart(partUri, sourcePart.ContentType);
            var destStream = destPart.GetStream(FileMode.Create, FileAccess.Write);
            var sourceStream = sourcePart.GetStream();
            var b = ((MemoryStream)sourceStream).GetBuffer();
            destStream.Write(b, 0, b.Length);
            destStream.Flush();
            return destPart;
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, string xml)
        {
            var destPart = CreatePart(partUri, contentType);
            var destStream = new StreamWriter(destPart.GetStream(FileMode.Create, FileAccess.Write));
            destStream.Write(xml);
            destStream.Flush();
            return destPart;
        }

        internal ZipPackagePart GetPart(Uri partUri)
        {
            if (PartExists(partUri))
            {
                return Parts.Single(x => x.Key.Equals(GetUriKey(partUri.OriginalString),StringComparison.OrdinalIgnoreCase)).Value;
            }
            else
            {
                throw (new InvalidOperationException("Part does not exist."));
            }
        }

        internal string GetUriKey(string uri)
        {
            string ret = uri.Replace('\\', '/');
            if (ret[0] != '/')
            {
                ret = '/' + ret;
            }
            return ret;
        }
        internal bool PartExists(Uri partUri)
        {
            var uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());
            return Parts.ContainsKey(uriKey);
        }
#endregion

        internal void DeletePart(Uri Uri)
        {
            var delList=new List<object[]>(); 
            foreach (var p in Parts.Values)
            {
                foreach (var r in p.GetRelationships())
                {                    
                    if (r.TargetUri !=null && UriHelper.ResolvePartUri(p.Uri, r.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
                    {                        
                        delList.Add(new object[]{r.Id, p});
                    }
                }
            }
            foreach (var o in delList)
            {
                ((ZipPackagePart)o[1]).DeleteRelationship(o[0].ToString());
            }
            var rels = GetPart(Uri).GetRelationships();
            while (rels.Count > 0)
            {
                rels.Remove(rels.First().Id);
            }
            rels=null;
            _contentTypes.Remove(GetUriKey(Uri.OriginalString));
            //remove all relations
            Parts.Remove(GetUriKey(Uri.OriginalString));
            
        }
        internal void Save(Stream stream)
        {
            var enc = Encoding.UTF8;
            ZipOutputStream os = new ZipOutputStream(stream, true);
            os.EnableZip64 = Zip64Option.AsNecessary;
            os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)_compression;            

            /**** ContentType****/
            var entry = os.PutNextEntry("[Content_Types].xml");
            byte[] b = enc.GetBytes(GetContentTypeXml());
            os.Write(b, 0, b.Length);
            /**** Top Rels ****/
            _rels.WriteZip(os, $"_rels/.rels");
            List<ZipPackagePart> saveAfterParts = new List<ZipPackagePart>();
            var partsToSave = Parts.Values.ToList();
            for(int i=0;i < partsToSave.Count;i++)            
            {
                var part = partsToSave[i];
                //Shared strings, metadata and rich data must be saved after the cells have been saved as references to are updated while saving the worksheets.
                if (part.ContentType == ContentTypes.contentTypeSharedString || 
                    part.ContentType == ContentTypes.contentTypeMetaData ||
                    part.ContentType == ContentTypes.contentTypeRichDataValueType ||
                    part.ContentType == ContentTypes.contentTypeRichDataValueStructure ||
                    part.ContentType == ContentTypes.contentTypeRichDataValue ||
                    part.ContentType == ContentTypes.contentTypeWorkbookDefault ||
                    part.ContentType == ContentTypes.contentTypeWorkbookMacroEnabled)
                {
                    saveAfterParts.Add(part);
                }
                else
                {
                    part.WriteZip(os);
                }
            }

            foreach (var part in saveAfterParts)
            {
                if (part.ShouldBeSaved==false)
                {
                    DeletePart(part.Uri);
                }
            }

            //Shared strings must be saved after all worksheets. The ss dictionary is populated when that workheets are saved (to get the best performance).
            foreach (var part in saveAfterParts)
            {
                part.WriteZip(os);
            }
            os.Flush();
            
            os.Close();
            os.Dispose();  
            
            //return ms;
        }

        private string GetContentTypeXml()
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (ContentType ct in _contentTypes.Values)
            {
                if (ct.IsExtension)
                {
                    xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
                }
                else
                {
                    var uriKey = GetUriKey(ct.Match);
                    if (Parts.TryGetValue(uriKey, out ZipPackagePart part) && part.ShouldBeSaved)
                    {
                        xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, uriKey);
                    }
                }
            }
            xml.Append("</Types>");
            return xml.ToString();
        }
        internal void Flush()
        {

        }
        internal void Close()
        {
            
        }

        public void Dispose()
        {
            foreach(var part in Parts.Values)
            {
                part.Dispose();
            }
            _zip?.Dispose();
        }

        CompressionLevel _compression = CompressionLevel.Default;
        /// <summary>
        /// Compression level
        /// </summary>
        public CompressionLevel Compression 
        { 
            get
            {
                return _compression;
            }
            set
            {
                foreach (var part in Parts.Values)
                {
                    if (part.CompressionLevel == _compression)
                    {
                        part.CompressionLevel = value;
                    }
                }
                _compression = value;
            }
        }
    }
}
