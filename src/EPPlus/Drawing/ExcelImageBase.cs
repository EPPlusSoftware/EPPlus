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

namespace OfficeOpenXml.Drawing
{
    
    /// <summary>
    /// Represents an read only image 
    /// </summary>
    public class ExcelImageReadOnly
    {
        internal ExcelImageReadOnly()
        {
                
        }
        internal IPictureContainer _container;
        internal ePictureType[] _restrictedTypes = new ePictureType[0];
        internal ExcelImageReadOnly(IPictureContainer container, ePictureType[] restrictedTypes = null)
        {
            _container = container;
            if (restrictedTypes != null)
            {
                _restrictedTypes = restrictedTypes;
            }
        }
        /// <summary>
        /// If this object contains an image.
        /// </summary>
        public bool HasImage
        {
            get
            {
                return Type.HasValue;
            }
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

        internal void SetImage(byte[] image, ePictureType pictureType, bool removePrevImage)
        {
            if (_container == null)
            {
                SetImageNoContainer(image, pictureType);
            }
            else
            {
                SetImageContainer(image, pictureType, removePrevImage);
            }
        }

        private void SetImageContainer(byte[] image, ePictureType pictureType, bool removePrevImage)
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

            using (var ms = RecyclableMemory.GetStream(image))
            {
                if (_container.RelationDocument.Package.Settings.ImageSettings.GetImageBounds(ms, pictureType, out double width, out double height, out double horizontalResolution, out double verticalResolution))
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
            }
            _container.SetNewImage();
        }

        internal void SetRestrictedTypes(ePictureType[] restrictedTypes)
        {
            _restrictedTypes = restrictedTypes;
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
        internal void SetImageNoContainer(byte[] image, ePictureType pictureType)
        {
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
                    ImageBytes = img;
                    pictureType = pt.Value;
                }
            }
            else
            {
                ImageBytes = image;
            }
            using (var ms = RecyclableMemory.GetStream(image))
            {
                var imageHandler = new GenericImageHandler();
                if (imageHandler.GetImageBounds(ms, pictureType, out double width, out double height, out double horizontalResolution, out double verticalResolution))
                {
                    Bounds.Width = width;
                    Bounds.Height = height;
                    Bounds.HorizontalResolution = horizontalResolution;
                    Bounds.VerticalResolution = verticalResolution;
                }
                else
                {
                    throw (new InvalidOperationException($"The image format is not supported: {pictureType} or the image is corrupt "));
                }
            }
        }

    }
}