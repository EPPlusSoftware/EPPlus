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
namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Type of color transformation.
    /// See OOXML documentation section 20.1.2.3 for more detailed information.
    /// </summary>
    public enum eColorTransformType
    {
        /// <summary>
        /// A lighter version of its input color.
        /// </summary>
        Tint,
        /// <summary>
        /// A darker version of its input color
        /// </summary>
        Shade,
        /// <summary>
        /// The color rendered should be the complement of its input color
        /// </summary>
        Comp,
        /// <summary>
        /// The inverse of its input color
        /// </summary>
        Inv,
        /// <summary>
        /// A grayscale of its input color
        /// </summary>
        Gray,
        /// <summary>
        /// Apply an opacity to the input color
        /// </summary>
        Alpha,
        /// <summary>
        /// Apply a more or less opaque version of the input color
        /// </summary>
        AlphaOff,
        /// <summary>
        /// The opacity as expressed by a percentage offset increase or decrease of the input color
        /// </summary>
        AlphaMod,
        /// <summary>
        /// Sets the hue
        /// </summary>
        Hue,
        /// <summary>
        /// The input color with its hue shifted
        /// </summary>
        HueOff,
        /// <summary>
        /// The input color with its hue modulated by the given percentage
        /// </summary>
        HueMod,
        /// <summary>
        /// Sets the saturation
        /// </summary>
        Sat,
        /// <summary>
        /// The saturation as expressed by a percentage offset increase or decrease of the input color
        /// </summary>
        SatOff,
        /// <summary>
        /// The saturation as expressed by a percentage relative to the input color
        /// </summary>
        SatMod,
        /// <summary>
        /// Sets the luminance
        /// </summary>
        Lum,
        /// <summary>
        /// The luminance as expressed by a percentage offset increase or decrease of the input color
        /// </summary>
        LumOff,
        /// <summary>
        /// The luminance as expressed by a percentage relative to the input color
        /// </summary>
        LumMod,
        /// <summary>
        /// Sets the red component
        /// </summary>
        Red,
        /// <summary>
        /// The red component as expressed by a percentage offset increase or decrease of the input color
        /// </summary>
        RedOff,
        /// <summary>
        /// The red component as expressed by a percentage relative to the input color
        /// </summary>
        RedMod,
        /// <summary>
        /// Sets the green component
        /// </summary>
        Green,
        /// <summary>
        /// The green component as expressed by a percentage offset increase or decrease of the input color
        /// </summary>
        GreenOff,
        /// <summary>
        /// The green component as expressed by a percentage relative to the input color
        /// </summary>
        GreenMod,
        /// <summary>
        /// Sets the blue component
        /// </summary>
        Blue,
        /// <summary>
        /// The blue component as expressed by a percentage offset increase or decrease to the input color
        /// </summary>
        BlueOff,
        /// <summary>
        /// The blue component as expressed by a percentage relative to the input color
        /// </summary>
        BlueMod,
        /// <summary>
        /// Gamma shift of the input color
        /// </summary>
        Gamma,
        /// <summary>
        /// Inverse gamma shift of the input color
        /// </summary>
        InvGamma
    }
}