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
    public class ExcelImage : ExcelImageReadOnly
    {
        internal ExcelImage(IPictureContainer container, ePictureType[] restrictedTypes = null) : base(container, restrictedTypes) 
        {
        }
        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        public ExcelImage() 
        {

        }
        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imagePath">A path to the image file to load</param>
        public ExcelImage(string imagePath) 
        {
            SetImage(imagePath);
        }
        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageFile">A FileInfo referencing the image file to load</param>
        public ExcelImage(FileInfo imageFile)
        {
            SetImage(imageFile);
        }
        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageStream">The stream containing the image</param>
        /// <param name="pictureType">The type of image loaded in the stream</param>
        public ExcelImage(Stream imageStream, ePictureType pictureType)
        {
            SetImage(imageStream, pictureType);
        }
        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageBytes">The image as a byte array</param>
        /// <param name="pictureType">The type of image loaded in the stream</param>
        public ExcelImage(byte[] imageBytes, ePictureType pictureType)
        {
            SetImage(imageBytes, pictureType);
        }
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
        /// <param name="imageBytes">The image as a byte array.</param>
        /// <param name="pictureType">The type of image.</param>
        public ExcelImage SetImage(byte[] imageBytes, ePictureType pictureType)
        {
            SetImage(imageBytes, pictureType, true);
            return this;
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="image">The image object to use.</param>
        /// <seealso cref="ExcelImage"/>
        public ExcelImage SetImage(ExcelImage image)
        {
            if(image.Type==null)
            {
                throw new ArgumentNullException("Image type must not be null");
            }
            SetImage(image.ImageBytes, image.Type.Value, true);
            return this;
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public ExcelImage SetImage(Stream imageStream, ePictureType pictureType)
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
                var r = imageStream.Read(byRet, 0, (int)imageStream.Length);

                SetImage(byRet, pictureType);
            }
            return this;
        }
#if !NET35
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public async Task<ExcelImage> SetImageAsync(Stream imageStream, ePictureType pictureType)
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
                var r = await imageStream.ReadAsync(byRet, 0, (int)imageStream.Length);

                SetImage(byRet, pictureType);
            }
            return this;
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imagePath">The path to the image file.</param>
        public async Task<ExcelImage> SetImageAsync(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath), "Image Path cannot be empty");
            }
            var fi = new FileInfo(imagePath);
            return await SetImageAsync(fi);
        }
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageFile">The image file.</param>
        public async Task<ExcelImage> SetImageAsync(FileInfo imageFile)
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
            var r = await fs.ReadAsync(b, 0, b.Length);
            SetImage(b, type, true);
            return this;
        }

#endif
    }
}

