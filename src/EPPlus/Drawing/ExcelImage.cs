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
using System;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
#if NETFULL
using System.Drawing;
#endif
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Represents an image 
    /// </summary>
    public class ExcelImage
    {
        IPictureContainer _container;
        ePictureType[] _restrictedTypes;
        internal ExcelImage(IPictureContainer container, ePictureType[] restrictedTypes=null)
        {
            _container = container;
            _restrictedTypes = restrictedTypes ?? new ePictureType[0];
        }

        /// <summary>
        /// The type of image.
        /// </summary>
        public ePictureType? Type
        {
            get;
            internal set;
        }

        /// <summary>
        /// The image as a byte array.
        /// </summary>
        public byte[] ImageBytes 
        { 
            get;
            internal set; 
        }
        /// <summary>
        /// The image bounds and resolution
        /// </summary>
        public ExcelImageInfo Bounds
        {
            get;            
            internal set;
        } = new ExcelImageInfo();
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imagePath">The path to the image file.</param>
        public void SetImage(string imagePath)
        {
            if(string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath),"Image Path cannot be empty");
            }
            var fi=new FileInfo(imagePath); 
            if(fi.Exists==false)
            {
                throw new FileNotFoundException(imagePath);
            }
            var type = PictureStore.GetPictureType(fi.Extension);
            SetImage(File.ReadAllBytes(imagePath), type, true);
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageFile">The image file.</param>
        public void SetImage(FileInfo imageFile)
        {
            if (imageFile==null)
            {
                throw new ArgumentNullException(nameof(imageFile), "ImageFile cannot be null");
            }

            if (imageFile.Exists == false)
            {
                throw new FileNotFoundException(imageFile.FullName);
            }
            var type = PictureStore.GetPictureType(imageFile.Extension);
            SetImage(File.ReadAllBytes(imageFile.FullName), type, true);
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="image">The image as a byte array.</param>
        /// <param name="pictureType">The type of image.</param>
        public void SetImage(byte[] image, ePictureType pictureType)
        {
            SetImage(image, pictureType, true);
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public void SetImage(Stream imageStream, ePictureType pictureType)
        {
            if(imageStream is MemoryStream ms)
            {
                SetImage(ms.ToArray(), pictureType, true);
            }
            else
            {
                if(imageStream.CanRead ==false || imageStream.CanSeek == false)
                {
                    throw (new ArgumentException("Stream must be readable and seekble", nameof(imageStream)));
                }
                var byRet = new byte[imageStream.Length];
                imageStream.Seek(0, SeekOrigin.Begin);
                imageStream.Read(byRet, 0, (int)imageStream.Length);

                SetImage(byRet, pictureType);
            }
        }
#if !NET35 && !NET40
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public async Task SetImageAsync(Stream imageStream, ePictureType pictureType)
        {
            if (imageStream is MemoryStream ms)
            {
                SetImage(ms.ToArray(), pictureType, true);
            }
            else
            {
                if (imageStream.CanRead == false || imageStream.CanSeek == false)
                {
                    throw (new ArgumentException("Stream must be readable and seekble", nameof(imageStream)));
                }
                var byRet = new byte[imageStream.Length];
                imageStream.Seek(0, SeekOrigin.Begin);
                await imageStream.ReadAsync(byRet, 0, (int)imageStream.Length);

                SetImage(byRet, pictureType);
            }
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imagePath">The path to the image file.</param>
        public async Task SetImageAsync(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath), "Image Path cannot be empty");
            }
            var fi = new FileInfo(imagePath);
            await SetImageAsync(fi);
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageFile">The image file.</param>
        public async Task SetImageAsync(FileInfo imageFile)
        {
            if (imageFile == null)
            {
                throw new ArgumentNullException(nameof(imageFile), "ImageFile cannot be null");
            }

            if (imageFile.Exists == false)
            {
                throw new FileNotFoundException(imageFile.FullName);
            }
            var type = PictureStore.GetPictureType(imageFile.Extension);
            var fs = imageFile.OpenRead();
            var b = new byte[fs.Length];
            await fs.ReadAsync(b, 0, b.Length);
            SetImage(b, type, true);
        }

#endif
        internal ePictureType SetImage(byte[] image, ePictureType pictureType, bool removePrevImage)
        {
            ValidatePictureType(pictureType);
            Type = pictureType;
            if (pictureType == ePictureType.Wmz ||
               pictureType == ePictureType.Emz)
            {
                var img = ImageReader.ExtractImage(image, out ePictureType? pt);
                if (pt.HasValue)
                {
                    throw new ArgumentException($"Image is not of type {pictureType}.", nameof(image));
                }
                else
                {
                    if (string.IsNullOrEmpty(_container.ImageHash) == false && removePrevImage)
                    {
                        RemoveImageContainer();
                    }
                    ImageBytes = img;
                    pictureType = pt.Value;
                }
            }
            else
            {
                if (removePrevImage && string.IsNullOrEmpty(_container.ImageHash) == false)
                {
                    RemoveImageContainer();
                }
                ImageBytes = image;
            }
            PictureStore.SavePicture(image, _container, pictureType);
            var ms = new MemoryStream(image);
            if(_container.RelationDocument.Package.Settings.ImageSettings.GetImageBounds(ms, pictureType, out double height, out double width, out double horizontalResolution, out double verticalResolution))
            {
                Bounds.Width = width;
                Bounds.Height = height;
                Bounds.HorizontalResolution = horizontalResolution;
                Bounds.VerticalResolution = verticalResolution;
            }
            else
            {
                throw (new InvalidOperationException($"Image format not supported or: {pictureType} or corrupt image"));
            }

            //#if (Core)
            //            GetImageInformation(image, pictureType);
            //#else
            //            if(pictureType == ePictureType.Ico ||
            //               pictureType == ePictureType.Svg ||
            //               pictureType == ePictureType.WebP)
            //              { 
            //                  GetImageInformation(image, pictureType);
            //              }
            //              else
            //              {
            //                    try
            //                    {
            //                        var ms=new MemoryStream(image);
            //                        var img = Image.FromStream(ms);
            //                        Bounds.Width = img.Width;
            //                        Bounds.Height = img.Height;
            //                        Bounds.HorizontalResolution = img.HorizontalResolution;
            //                        Bounds.VerticalResolution = img.VerticalResolution;
            //                    }
            //                    catch
            //                    {
            //                        GetImageInformation(image, pictureType);
            //                    }                
            //               }
            //#endif
            _container.SetNewImage();
            return pictureType;
        }

        private void ValidatePictureType(ePictureType pictureType)
        {
            if (Array.Exists(_restrictedTypes, x => x == pictureType))
            {
                throw new InvalidOperationException($"Picture type {pictureType} is not supported for this operation.");
            }
        }

        internal void RemoveImage()
        {
            RemoveImageContainer();
            ImageBytes = null;
            Type = null;
            Bounds = new ExcelImageInfo();
        }
        private void RemoveImageContainer()
        {
            _container.RemoveImage();
            _container.RelPic = null;
            _container.ImageHash = null;
            _container.UriPic = null;
        }

        //private bool GetImageInformation(byte[] image, ePictureType pictureType)
        //{           
        //    double w = 0, h = 0;
        //    if (ImageReader.TryGetImageBounds(pictureType, new MemoryStream(image), ref w, ref h, out double hr, out double vr))
        //    {
        //        Bounds.Width = w;
        //        Bounds.Height = h;
        //        Bounds.HorizontalResolution = hr;
        //        Bounds.VerticalResolution = vr;
        //        return true;
        //    }
        //    return false;
        //}
    }
}

