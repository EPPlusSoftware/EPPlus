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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// The type of image a stream contains.
    /// </summary>
    public enum ePictureType
    {
        /// <summary>
        /// A bitmap image
        /// </summary>
        Bmp,
        /// <summary>
        /// A jpeg image
        /// </summary>
        Jpg,
        /// <summary>
        /// A gif image
        /// </summary>
        Gif,
        /// <summary>
        /// A png image
        /// </summary>
        Png,
        /// <summary>
        /// An Enhanced MetaFile image
        /// </summary>
        Emf,
        /// <summary>
        /// A Tiff image
        /// </summary>
        Tif,
        /// <summary>
        /// A windows metafile image
        /// </summary>
        Wmf,
        /// <summary>
        /// A Svg image
        /// </summary>
        Svg,
        /// <summary>
        /// A WebP image
        /// </summary>
        WebP,
        /// <summary>
        /// A Windows icon
        /// </summary>
        Ico,
        /// <summary>
        /// A compressed Enhanced MetaFile image
        /// </summary>
        Emz,
        /// <summary>
        /// A compressed Windows MetaFile image
        /// </summary>
        Wmz
    }
}
