using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Interfaces.Drawing.Image
{
    /// <summary>
    /// Returns bounds information about an image to EPPlus.
    /// </summary>
    public interface IImageHandler
    {
        /// <summary>
        /// Should return true if the text measurer is valid for this environment. 
        /// </summary>
        /// <returns>True if the measurer can be used else false.</returns>
        bool ValidForEnvironment();
        /// <summary>
        /// All types supported by the handler
        /// </summary>
        HashSet<ePictureType> SupportedTypes { get;  }
        /// <summary>
        /// Returns the boundrys and resolution of an image to EPPlus.
        /// </summary>
        /// <param name="image">The image stream</param>
        /// <param name="type">The type of image</param>
        /// <param name="width">The width returned. Must be larger than zero</param>
        /// <param name="height">The height returned. Must be larger than zero</param>
        /// <param name="horizontalResolution">Horizontal resolution. 96 is default.</param>
        /// <param name="verticalResolution">Vertical resolution. 96 is default.</param>
        /// <returns>Returns true if the operation succeeded, else false</returns>
        bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution);
        /// <summary>
        /// Returns the last exception in the GetImageBounds method.
        /// </summary>
        Exception LastException { get;  }
    }
}
