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
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Packaging
{
    internal class ZipPackagePart : ZipPackagePartBase, IDisposable
    {
        internal delegate void SaveHandlerDelegate(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName);

        internal ZipPackagePart(ZipPackage package, ZipEntry entry)
        {
            Package = package;
            Entry = entry;
            SaveHandler = null;
            Uri = new Uri(package.GetUriKey(entry.FileName), UriKind.Relative);
        }
        internal ZipPackagePart(ZipPackage package, Uri partUri, string contentType, CompressionLevel compressionLevel)
        {
            Package = package;
            Uri = partUri;
            ContentType = contentType;
            CompressionLevel = compressionLevel;
        }
        internal ZipPackage Package { get; set; }
        internal ZipEntry Entry { get; set; }
        internal CompressionLevel CompressionLevel;
        Stream _stream = null;
        internal Stream Stream
        {
            get
            {
                return _stream;
            }
            set
            {
                _stream = value;
            }
        }
        internal override ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
        {

            var rel = base.CreateRelationship(targetUri, targetMode, relationshipType);
            rel.SourceUri = Uri;
            return rel;
        }
        internal override ZipPackageRelationship CreateRelationship(string target, TargetMode targetMode, string relationshipType)
        {

            var rel = base.CreateRelationship(target, targetMode, relationshipType);
            rel.SourceUri = Uri;
            return rel;
        }

        internal Stream GetStream()
        {
            return GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
        }
        internal Stream GetStream(FileMode fileMode)
        {
            return GetStream(FileMode.Create, FileAccess.ReadWrite);
        }
        internal Stream GetStream(FileMode fileMode, FileAccess fileAccess)
        {
            if (_stream == null || fileMode == FileMode.CreateNew || fileMode == FileMode.Create)
            {
                _stream = RecyclableMemory.GetStream();
            }
            else
            {
                _stream.Seek(0, SeekOrigin.Begin);
            }
            return _stream;
        }

        string _contentType = "";
        public string ContentType
        {
            get
            {
                return _contentType;
            }
            internal set
            {
                if (!string.IsNullOrEmpty(_contentType))
                {
                    if (Package._contentTypes.ContainsKey(Package.GetUriKey(Uri.OriginalString)))
                    {
                        Package._contentTypes.Remove(Package.GetUriKey(Uri.OriginalString));
                        Package._contentTypes.Add(Package.GetUriKey(Uri.OriginalString), new ZipPackage.ContentType(value, false, Uri.OriginalString));
                    }
                }
                _contentType = value;
            }
        }
        public Uri Uri { get; private set; }
        public Stream GetZipStream()
        {
            MemoryStream ms = new MemoryStream();
            ZipOutputStream os = new ZipOutputStream(ms);
            return os;
        }
        internal SaveHandlerDelegate SaveHandler
        {
            get;
            set;
        }
        internal void WriteZip(ZipOutputStream os)
        {
            byte[] b;
            if (SaveHandler == null)
            {
                b = ((MemoryStream)GetStream()).ToArray();
                if (b.Length == 0)   //Make sure the file isn't empty. DotNetZip streams does not seems to handle zero sized files.
                {
                    return;
                }
                os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)CompressionLevel;
                os.PutNextEntry(Uri.OriginalString);
                os.Write(b, 0, b.Length);
            }
            else
            {
                SaveHandler(os, (CompressionLevel)CompressionLevel, Uri.OriginalString);
            }

            if (_rels.Count > 0)
            {
                string f = Uri.OriginalString;
                var name = Path.GetFileName(f);
                _rels.WriteZip(os, (string.Format("{0}_rels/{1}.rels", f.Substring(0, f.Length - name.Length), name)));
            }
            b = null;
        }


        public void Dispose()
        {
            _stream.Close();
            _stream.Dispose();
        }
    }
}
