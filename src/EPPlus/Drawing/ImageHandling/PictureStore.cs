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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.Packaging;
using System.Linq;
namespace OfficeOpenXml.Drawing
{
    internal class ImageInfo
    {
        internal string Hash { get; set; }
        internal Uri Uri { get; set; }
        internal int RefCount { get; set; }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal ExcelImageInfo Bounds { get; set; }
    }
    internal class PictureStore : IDisposable
    {
        ExcelPackage _pck;
        internal static int _id = 1;
        internal Dictionary<string, ImageInfo> _images;
        public PictureStore(ExcelPackage pck)
        {
            _pck = pck;
            _images = _pck.Workbook._images;
        }
        internal ImageInfo AddImage(byte[] image)
        {
            return AddImage(image, null, null);
        }
        internal ImageInfo AddImage(byte[] image, Uri uri, ePictureType? pictureType)
        {
            if (pictureType.HasValue == false) pictureType = ePictureType.Jpg;
#if (Core)
            var hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
            var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
            lock (_images)
            {
                if (_images.ContainsKey(hash))
                {
                    _images[hash].RefCount++;
                }
                else
                {
                    Packaging.ZipPackagePart imagePart;
                    string contentType;
                    if (uri == null)
                    {
                        var extension = GetExtension(pictureType.Value);
                        contentType = GetContentType(extension);
                        uri = GetNewUri(_pck.ZipPackage, "/xl/media/image{0}." + extension);
                        imagePart = _pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                        SaveImageToPart(image, imagePart);
                    }
                    else
                    {
                        var extension = GetExtension(uri);
                        contentType = GetContentType(extension);
                        if (pictureType == null)
                        {
                            pictureType = GetPictureType(extension);
                        }
                        if (_pck.ZipPackage.PartExists(uri))
                        {
                            if(_images.Values.Any(x=>x.Uri.OriginalString==uri.OriginalString))
                            {
                                uri = GetNewUri(_pck.ZipPackage, "/xl/media/image{0}." + extension);
                                imagePart = _pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                                SaveImageToPart(image, imagePart);
                            }
                            else
                            {
                                imagePart = _pck.ZipPackage.GetPart(uri);
                            }
                        }
                        else
                        {
                            imagePart = _pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                            SaveImageToPart(image, imagePart);
                        }
                    }
                    _images.Add(hash,
                        new ImageInfo()
                        {
                            Uri = uri,
                            RefCount = 1,
                            Hash = hash,
                            Part = imagePart,
                            Bounds = GetImageBounds(image, pictureType.Value, _pck)
                        });
                }
            }
            return _images[hash];
        }

        private static void SaveImageToPart(byte[] image, ZipPackagePart imagePart)
        {
            var stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
            stream.Write(image, 0, image.GetLength(0));
            stream.Flush();
        }

        internal static ExcelImageInfo GetImageBounds(byte[] image, ePictureType type, ExcelPackage pck)
        {
            var ret = new ExcelImageInfo();
            var ms = new MemoryStream(image);
            var s = pck.Settings.ImageSettings;

            if(s.GetImageBounds(ms, type, out double width, out double height, out double horizontalResolution, out double verticalResolution)==false)
            {
                throw (new InvalidOperationException($"No image handler for image type {type}"));
            }
            ret.Width = width;
            ret.Height = height;
            ret.HorizontalResolution = horizontalResolution;
            ret.VerticalResolution = verticalResolution;
            return ret;
        }
        internal static string GetExtension(Uri uri)
        {
            var s = uri.OriginalString;
            var i = s.LastIndexOf('.');
            if(i>0)
            {
                return s.Substring(i + 1);
            }
            return null;
        }

        internal ImageInfo LoadImage(byte[] image, Uri uri, Packaging.ZipPackagePart imagePart)
        {
#if (Core)
            var hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
            var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
            if (_images.ContainsKey(hash))
            {
                _images[hash].RefCount++;
            }
            else
            {
                _images.Add(hash, new ImageInfo() { Uri = uri, RefCount = 1, Hash = hash, Part = imagePart });
            }
            return _images[hash];
        }
        internal void RemoveImage(string hash, IPictureContainer container)
        {
            lock (_images)
            {
                if (_images.ContainsKey(hash))
                {
                    var ii = _images[hash];
                    ii.RefCount--;
                    if (ii.RefCount == 0)
                    {
                        _pck.ZipPackage.DeletePart(ii.Uri);
                        _images.Remove(hash);
                    }
                }
                if(container.RelationDocument.Hashes.ContainsKey(hash))
                {
                    container.RelationDocument.Hashes[hash].RefCount--;
                    if (container.RelationDocument.Hashes[hash].RefCount <= 0)
                    {
                        container.RelationDocument.Hashes.Remove(hash);
                    }
                        
                }
            }
        }
        internal ImageInfo GetImageInfo(byte[] image)
        {
            var hash = GetHash(image);
            if (_images.ContainsKey(hash))
            {
                return _images[hash];
            }
            else
            {
                return null;
            }
        }
        internal bool ImageExists(byte[] image)
        {
            var hash = GetHash(image);
            return _images.ContainsKey(hash);
        }

        internal static string GetHash(byte[] image)
        {
#if (Core)
            var hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
            return BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
        }

        private Uri GetNewUri(Packaging.ZipPackage package, string sUri)
        {
            Uri uri;
            do
            {
                uri = new Uri(string.Format(sUri, _id++), UriKind.Relative);
            }
            while (package.PartExists(uri));
            return uri;
        }
        
        internal static byte[] GetPicture(string relId, IPictureContainer container, out string contentType, out ePictureType pictureType)
        {
            ZipPackagePart part;
            container.RelPic = container.RelationDocument.RelatedPart.GetRelationship(relId);
            container.UriPic = UriHelper.ResolvePartUri(container.RelationDocument.RelatedUri, container.RelPic.TargetUri);
            part = container.RelationDocument.RelatedPart.Package.GetPart(container.UriPic);

            var extension = Path.GetExtension(container.UriPic.OriginalString);
            contentType = GetContentType(extension);
            pictureType = GetPictureType(extension);
            return ((MemoryStream)part.GetStream()).ToArray();
        }
        internal static ePictureType GetPictureType(Uri uri)
        {
            var ext = GetExtension(uri);
            return GetPictureType(ext);
        }
        internal static ePictureType GetPictureType(string extension)
        {
            if (extension.StartsWith(".", StringComparison.OrdinalIgnoreCase))
                extension = extension.Substring(1);

            switch (extension.ToLower(CultureInfo.InvariantCulture))
            {
                case "bmp":
                case "dib":
                    return ePictureType.Bmp;
                case "jpg":
                case "jpeg":
                case "jfif":
                case "jpe":
                case "exif":
                    return ePictureType.Jpg;
                case "gif":
                    return ePictureType.Gif;
                case "png":
                    return ePictureType.Png;
                case "emf":
                    return ePictureType.Emf;
                case "emz":
                    return ePictureType.Emz;
                case "tif":
                case "tiff":
                    return ePictureType.Tif;
                case "wmf":
                    return ePictureType.Wmf;
                case "wmz":
                    return ePictureType.Wmz;
                case "webp":
                    return ePictureType.WebP;
                case "ico":
                    return ePictureType.Ico;
                case "svg":
                    return ePictureType.Svg;
                default:
                    throw (new InvalidOperationException($"Image with extension {extension} is not supported."));
            }
        }
        internal static string GetExtension(ePictureType type)
        {
            switch (type)
            {
                case ePictureType.Bmp:
                    return "bmp";
                case ePictureType.Gif:
                    return "gif";
                case ePictureType.Png:
                    return "png";
                case ePictureType.Emf:
                    return "emf";
                case ePictureType.Wmf:
                    return "wmf";
                case ePictureType.Tif:
                    return "tif";
                case ePictureType.WebP:
                    return "webp";
                case ePictureType.Ico:
                    return "ico";
                case ePictureType.Svg:
                    return "svg";
                default:
                    return "jpg";
            }
        }

        internal static string GetContentType(string extension)
        {
            if (extension.StartsWith(".", StringComparison.OrdinalIgnoreCase))
                extension = extension.Substring(1);

            switch (extension.ToLower(CultureInfo.InvariantCulture))
            {
                case "bmp":
                case "dib":
                    return "image/bmp";
                case "jpg":
                case "jpeg":
                case "jfif":
                case "jpe":
                    return "image/jpeg";
                case "gif":
                    return "image/gif";
                case "png":
                    return "image/png";
                case "cgm":
                    return "image/cgm";
                case "emf":
                case "emz":
                    return "image/x-emf";
                case "eps":
                    return "image/x-eps";
                case "pcx":
                    return "image/x-pcx";
                case "tga":
                    return "image/x-tga";
                case "tif":
                case "tiff":
                    return "image/x-tiff";
                case "wmf":
                case "wmz":
                    return "image/x-wmf";
                case "svg":
                    return "image/svg+xml";
                case "webp":
                    return "image/webp";
                case "ico":
                    return "image/x-icon";
                default:
                    return "image/jpeg";
            }
        }
        internal static string SavePicture(byte[] image, IPictureContainer container, ePictureType type)
        {
            var store = container.RelationDocument.Package.PictureStore;
            var ii = store.AddImage(image, container.UriPic, type);

            container.ImageHash = ii.Hash;
            var hashes = container.RelationDocument.Hashes;
            if (hashes.ContainsKey(ii.Hash))
            {
                var relID = hashes[ii.Hash].RelId;
                container.RelPic = container.RelationDocument.RelatedPart.GetRelationship(relID);
                container.UriPic = UriHelper.ResolvePartUri(container.RelPic.SourceUri, container.RelPic.TargetUri);
                return relID;
            }
            else
            {
                container.UriPic = ii.Uri;
                container.ImageHash = ii.Hash;
            }

            //Set the Image and save it to the package.
            container.RelPic = container.RelationDocument.RelatedPart.CreateRelationship(UriHelper.GetRelativeUri(container.RelationDocument.RelatedUri, container.UriPic), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            //AddNewPicture(img, picRelation.Id);
            hashes.Add(ii.Hash, new HashInfo(container.RelPic.Id) { RefCount = 1});

            return container.RelPic.Id;
        }

        public void Dispose()
        {
            _images = null;
        }
    }
}
