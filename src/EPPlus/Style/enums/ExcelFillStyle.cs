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
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Fill pattern
    /// </summary>
    public enum ExcelFillStyle
    {
        /// <summary>
        /// No fill
        /// </summary>
        None,
        /// <summary>
        /// A solid fill
        /// </summary>
        Solid,
        /// <summary>
        /// Dark gray  <para/>
        /// Excel name: 75% Gray
        /// </summary>
        DarkGray,
        /// <summary>
        /// Medium gray <para/>
        /// Excel name: 50% Gray
        /// </summary>
        MediumGray,
        /// <summary>
        /// Light gray <para/>
        /// Excel name: 25% Gray
        /// </summary>
        LightGray,
        /// <summary>
        /// Grayscale of 0.125, 1/8 <para/>
        /// Excel name: 12.5% Gray
        /// </summary>
        Gray125,
        /// <summary>
        /// Grayscale of 0.0625, 1/16 <para/>
        /// Excel name: 6.25% Gray
        /// </summary>
        Gray0625,
        /// <summary>
        /// Dark vertical <para/>
        /// Excel name: Vertical Stripe
        /// </summary>
        DarkVertical,
        /// <summary>
        /// Dark horizontal <para/>
        /// Excel name: Horizontal Stripe
        /// </summary>
        DarkHorizontal,
        /// <summary>
        /// Dark down <para/>
        /// Excel name: Reverse Diagonal Stripe
        /// </summary>
        DarkDown,
        /// <summary>
        /// Dark up <para/>
        /// Excel name: Diagonal Stripe
        /// </summary>
        DarkUp,
        /// <summary>
        /// Dark grid <para/>
        /// Excel name: Diagonal Crosshatch
        /// </summary>
        DarkGrid,
        /// <summary>
        /// Dark trellis <para/>
        /// Excel name: Thick Diagonal Crosshatch
        /// </summary>
        DarkTrellis,
        /// <summary>
        /// Light vertical <para/>
        /// Excel name: Thin Vertical Stripe
        /// </summary>
        LightVertical,
        /// <summary>
        /// Light horizontal <para/>
        /// Excel name: Thin Horizontal Stripe
        /// </summary>
        LightHorizontal,
        /// <summary>
        /// Light down <para/>
        /// Excel name: Thin Reverse Diagonal Stripe
        /// </summary>
        LightDown,
        /// <summary>
        /// Light up <para/>
        /// Excel name: Thin Diagonal Stripe
        /// </summary>
        LightUp,
        /// <summary>
        /// Light grid <para/>
        /// Excel name: Thin Horizontal Crosshatch
        /// </summary>
        LightGrid,
        /// <summary>
        /// Light trellis <para/>
        /// Excel name: Thin Diagonal Crosshatch
        /// </summary>
        LightTrellis
    }
}