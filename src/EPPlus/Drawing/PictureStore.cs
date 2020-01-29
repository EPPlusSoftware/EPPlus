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
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing
{
    internal class PictureStore : IDisposable
    {
        ExcelPackage _pck;
        internal static int _id = 1;
        internal Dictionary<string, ImageInfo> _images = new Dictionary<string, ImageInfo>();
        public PictureStore(ExcelPackage pck)
        {
            _pck = pck;
        }
        internal class ImageInfo
        {
            internal string Hash { get; set; }
            internal Uri Uri { get; set; }
            internal int RefCount { get; set; }
            internal Packaging.ZipPackagePart Part { get; set; }
        }
        internal ImageInfo AddImage(byte[] image)
        {
            return AddImage(image, null, "");
        }
        internal ImageInfo AddImage(byte[] image, Uri uri, string contentType)
        {
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
                    if (uri == null)
                    {
                        uri = GetNewUri(_pck.Package, "/xl/media/image{0}.jpg");
                        imagePart = _pck.Package.CreatePart(uri, "image/jpeg", CompressionLevel.None);
                    }
                    else
                    {
                        imagePart = _pck.Package.CreatePart(uri, contentType, CompressionLevel.None);
                    }
                    var stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
                    stream.Write(image, 0, image.GetLength(0));

                    _images.Add(hash, new ImageInfo() { Uri = uri, RefCount = 1, Hash = hash, Part = imagePart });
                }
            }
            return _images[hash];
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
                        _pck.Package.DeletePart(ii.Uri);
                        _images.Remove(hash);
                    }
                }
                if(container.RelationDocument.Hashes.ContainsKey(hash))
                {
                    container.RelationDocument.Hashes[hash].RefCount--;
                    if (container.RelationDocument.Hashes[hash].RefCount == 0)
                    {
                        container.RelationDocument.Hashes.Remove(hash);
                    }
                        
                }
            }
        }
        internal ImageInfo GetImageInfo(byte[] image)
        {
#if (Core)
            var hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
            var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");

            if (_images.ContainsKey(hash))
            {
                return _images[hash];
            }
            else
            {
                return null;
            }
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

        internal static Image GetPicture(string relId, IPictureContainer container, out string contentType)
        {
            ZipPackagePart part;
            //if (container.Drawing is ExcelChart chart)
            //{
            //    container.RelPic = chart.Part.GetRelationship(relId);
            //    container.UriPic = UriHelper.ResolvePartUri(chart.UriChart, container.RelPic.TargetUri);
            //    part = chart.Part.Package.GetPart(container.UriPic);
            //}
            //else
            //{                
                container.RelPic = container.RelationDocument.RelatedPart.GetRelationship(relId);
                container.UriPic = UriHelper.ResolvePartUri(container.RelationDocument.RelatedUri, container.RelPic.TargetUri);
                part = container.RelationDocument.RelatedPart.Package.GetPart(container.UriPic);
            //}
            
            FileInfo f = new FileInfo(container.UriPic.OriginalString);
            contentType = PictureStore.GetContentType(f.Extension);
            return Image.FromStream(part.GetStream());
        }

        internal static string GetContentType(string extension)
        {
            switch (extension.ToLower(CultureInfo.InvariantCulture))
            {
                case ".bmp":
                    return "image/bmp";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".png":
                    return "image/png";
                case ".cgm":
                    return "image/cgm";
                case ".emf":
                    return "image/x-emf";
                case ".eps":
                    return "image/x-eps";
                case ".pcx":
                    return "image/x-pcx";
                case ".tga":
                    return "image/x-tga";
                case ".tif":
                case ".tiff":
                    return "image/x-tiff";
                case ".wmf":
                    return "image/x-wmf";
                default:
                    return "image/jpeg";

            }
        }
        internal static ImageFormat GetImageFormat(string contentType)
        {
            switch (contentType.ToLower(CultureInfo.InvariantCulture))
            {
                case "image/bmp":
                    return ImageFormat.Bmp;
                case "image/jpeg":
                    return ImageFormat.Jpeg;
                case "image/gif":
                    return ImageFormat.Gif;
                case "image/png":
                    return ImageFormat.Png;
                case "image/x-emf":
                    return ImageFormat.Emf;
                case "image/x-tiff":
                    return ImageFormat.Tiff;
                case "image/x-wmf":
                    return ImageFormat.Wmf;
                default:
                    return ImageFormat.Jpeg;

            }
        }        //Add a new image to the compare collection
        internal static string SavePicture(Image image, IPictureContainer container)
        {
#if (Core)
            byte[] img = ImageCompat.GetImageAsByteArray(image);
#else
            ImageConverter ic = new ImageConverter();
            byte[] img = (byte[])ic.ConvertTo(image, typeof(byte[]));
#endif
            var store = container.RelationDocument.Package.PictureStore;
            var ii = store.AddImage(img);

            container.ImageHash = ii.Hash;
            var hashes = container.RelationDocument.Hashes;
            if (hashes.ContainsKey(ii.Hash))
            {
                var relID = hashes[ii.Hash].RelId;
                var rel = container.RelationDocument.RelatedPart.GetRelationship(relID);
                container.UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
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
            hashes.Add(ii.Hash, new HashInfo(container.RelPic.Id));

            return container.RelPic.Id;
        }

        public void Dispose()
        {
            _images = null;
        }
    }
}
