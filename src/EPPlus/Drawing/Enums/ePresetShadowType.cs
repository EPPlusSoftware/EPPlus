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
    /// This enum indicates one of 20 preset OOXML shadow types.
    /// This values does NOT correspond to the the preset types in Excel. 
    /// Please use the SetPresetShadow method for Excel preset types.
    /// <seealso cref="Style.Effect.ExcelDrawingEffectStyle.SetPresetShadow"/>
    /// </summary>
    public enum ePresetShadowType
    {
        /// <summary>
        /// 1. Top Left Drop Shadow, Default
        /// </summary>
        TopLeftDropShadow,
        /// <summary>
        /// 2. Top Right Drop Shadow
        /// </summary>
        TopRightDropShadow,
        /// <summary>
        /// 3.
        /// </summary>
        BackLeftPerspectiveShadow,
        /// <summary>
        /// 4. Back Right Perspective Shadow
        /// </summary>
        BackRightPerspectiveShadow,
        /// <summary>
        /// 5. Bottom Left Drop Shadow
        /// </summary>
        BottomLeftDropShadow,
        /// <summary>
        /// 6. Bottom Right Drop Shadow
        /// </summary>
        BottomRightDropShadow,
        /// <summary>
        /// 7. FrontLeftPerspectiveShadow
        /// </summary>
        FrontLeftPerspectiveShadow,
        /// <summary>
        /// 8. Front Right Perspective Shadow
        /// </summary>
        FrontRightPerspectiveShadow,
        /// <summary>
        /// 9. Top Left Small DropShadow
        /// </summary>
        TopLeftSmallDropShadow,
        /// <summary>
        /// 10. Top Left Large Drop Shadow
        /// </summary>
        TopLeftLargeDropShadow,
        /// <summary>
        /// 11. Back Left Long Perspective Shadow
        /// </summary>
        BackLeftLongPerspectiveShadow,
        /// <summary>
        /// Back Right Long Perspective Shadow
        /// </summary>
        BackRightLongPerspectiveShadow,
        /// <summary>
        /// 13. Top Left Double Drop Shadow
        /// </summary>
        TopLeftDoubleDropShadow,
        /// <summary>
        /// 14. Bottom Right Small Drop Shadow
        /// </summary>
        BottomRightSmallDropShadow,
        /// <summary>
        /// 15. Front Left Long Perspective Shadow
        /// </summary>
        FrontLeftLongPerspectiveShadow,
        /// <summary>
        /// 16. Front Right LongPerspective Shadow
        /// </summary>
        FrontRightLongPerspectiveShadow,
        /// <summary>
        /// 17.  3D Outer Box Shadow
        /// </summary>
        OuterBoxShadow3D,
        /// <summary>
        /// 18. 3D Inner Box Shadow
        /// </summary>
        InnerBoxShadow3D,
        /// <summary>
        /// 19. Back Center Perspective Shadow
        /// </summary>
        BackCenterPerspectiveShadow,
        /// <summary>
        /// 20. Front Bottom Shadow
        /// </summary>
        FrontBottomShadow
    }
}