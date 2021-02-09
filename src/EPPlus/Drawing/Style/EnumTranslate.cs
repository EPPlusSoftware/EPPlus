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
using OfficeOpenXml.Utils.Extensions;
using System;

namespace OfficeOpenXml.Drawing.Style
{
    /// <summary>
    /// This class contains translation between enums and the actual xml values.
    /// </summary>
    internal static class EnumTranslateExtentions
    {
        #region eRectangleAlignment
        internal static string TranslateString(this eRectangleAlignment alignment)
        {
            switch (alignment)
            {
                case eRectangleAlignment.BottomLeft:
                    return "bl";
                case eRectangleAlignment.BottomRight:
                    return "br";
                case eRectangleAlignment.Center:
                    return "ctr";
                case eRectangleAlignment.Left:
                    return "l";
                case eRectangleAlignment.Right:
                    return "r";
                case eRectangleAlignment.Top:
                    return "t";
                case eRectangleAlignment.TopLeft:
                    return "tl";
                case eRectangleAlignment.TopRight:
                    return "tr";
                default:
                    return "b";
            }
        }
        internal static eRectangleAlignment TranslateRectangleAlignment(this string v)
        {
            switch (v.ToLower())
            {
                case "bl":
                    return eRectangleAlignment.BottomLeft;
                case "br":
                    return eRectangleAlignment.BottomRight;
                case "ctr":
                    return eRectangleAlignment.Center;
                case "l":
                    return eRectangleAlignment.Left;
                case "r":
                    return eRectangleAlignment.Right;
                case "t":
                    return eRectangleAlignment.Top;
                case "tl":
                    return eRectangleAlignment.TopLeft;
                case "tr":
                    return eRectangleAlignment.TopRight;
                default:
                    return eRectangleAlignment.Bottom;
            }
        }
        #endregion
        #region ePresetShadowType
        internal static string TranslateString(this ePresetShadowType v)
        {
            switch (v)
            {
                case ePresetShadowType.TopLeftDropShadow:
                    return "shdw1";
                case ePresetShadowType.TopRightDropShadow:
                    return "shdw2";
                case ePresetShadowType.BackLeftPerspectiveShadow:
                    return "shdw3";
                case ePresetShadowType.BackRightPerspectiveShadow:
                    return "shdw4";
                case ePresetShadowType.BottomLeftDropShadow:
                    return "shdw5";
                case ePresetShadowType.BottomRightDropShadow:
                    return "shdw6";
                case ePresetShadowType.FrontLeftPerspectiveShadow:
                    return "shdw7";
                case ePresetShadowType.FrontRightPerspectiveShadow:
                    return "shdw8";
                case ePresetShadowType.TopLeftSmallDropShadow:
                    return "shdw9";
                case ePresetShadowType.TopLeftLargeDropShadow:
                    return "shdw10";
                case ePresetShadowType.BackLeftLongPerspectiveShadow:
                    return "shdw11";
                case ePresetShadowType.BackRightLongPerspectiveShadow:
                    return "shdw12";
                case ePresetShadowType.TopLeftDoubleDropShadow:
                    return "shdw13";
                case ePresetShadowType.BottomRightSmallDropShadow:
                    return "shdw14";
                case ePresetShadowType.FrontLeftLongPerspectiveShadow:
                    return "shdw15";
                case ePresetShadowType.FrontRightLongPerspectiveShadow:
                    return "shdw16";
                case ePresetShadowType.OuterBoxShadow3D:
                    return "shdw17";
                case ePresetShadowType.InnerBoxShadow3D:
                    return "shdw18";
                case ePresetShadowType.BackCenterPerspectiveShadow:
                    return "shdw19";
                case ePresetShadowType.FrontBottomShadow:
                    return "shdw20";
                default:
                    return "shdw1";
            }
        }

        internal static ePresetShadowType TranslatePresetShadowType(this string s)
        {
            switch (s)
            {
                case "shdw1":
                    return ePresetShadowType.TopLeftDropShadow;
                case "shdw2":
                    return ePresetShadowType.TopRightDropShadow;
                case "shdw3":
                    return ePresetShadowType.BackLeftPerspectiveShadow;
                case "shdw4":
                    return ePresetShadowType.BackRightPerspectiveShadow;
                case "shdw5":
                    return ePresetShadowType.BottomLeftDropShadow;
                case "shdw6":
                    return ePresetShadowType.BottomRightDropShadow;
                case "shdw7":
                    return ePresetShadowType.FrontLeftPerspectiveShadow;
                case "shdw8":
                    return ePresetShadowType.FrontRightPerspectiveShadow;
                case "shdw9":
                    return ePresetShadowType.TopLeftSmallDropShadow;
                case "shdw10":
                    return ePresetShadowType.TopLeftLargeDropShadow;
                case "shdw11":
                    return ePresetShadowType.BackLeftLongPerspectiveShadow;
                case "shdw12":
                    return ePresetShadowType.BackRightLongPerspectiveShadow;
                case "shdw13":
                    return ePresetShadowType.TopLeftDoubleDropShadow;
                case "shdw14":
                    return ePresetShadowType.BottomRightSmallDropShadow;
                case "shdw15":
                    return ePresetShadowType.FrontLeftLongPerspectiveShadow;
                case "shdw16":
                    return ePresetShadowType.FrontRightLongPerspectiveShadow;
                case "shdw17":
                    return ePresetShadowType.OuterBoxShadow3D;
                case "shdw18":
                    return ePresetShadowType.InnerBoxShadow3D;
                case "shdw19":
                    return ePresetShadowType.BackCenterPerspectiveShadow;
                case "shdw20":
                    return ePresetShadowType.FrontBottomShadow;
                default:
                    return ePresetShadowType.TopLeftDropShadow;
            }
        }
        #endregion
        #region eLightRigDirection
        internal static eLightRigDirection TranslateLightRigDirection(this string s)
        {
            switch (s)
            {
                case "b":
                    return eLightRigDirection.Bottom;
                case "bl":
                    return eLightRigDirection.BottomLeft;
                case "br":
                    return eLightRigDirection.BottomRight;
                case "l":
                    return eLightRigDirection.Left;
                case "r":
                    return eLightRigDirection.Right;
                case "t":
                    return eLightRigDirection.Top;
                case "tl":
                    return eLightRigDirection.TopLeft;
                case "tr":
                    return eLightRigDirection.TopRight;
                default:
                    return eLightRigDirection.Bottom;
            }
        }
        internal static string TranslateString(this eLightRigDirection v)
        {
            switch (v)
            {
                case eLightRigDirection.Bottom:
                    return "b";
                case eLightRigDirection.BottomLeft:
                    return "bl";
                case eLightRigDirection.BottomRight:
                    return "br";
                case eLightRigDirection.Left:
                    return "l";
                case eLightRigDirection.Right:
                    return "r";
                case eLightRigDirection.Top:
                    return "t";
                case eLightRigDirection.TopLeft:
                    return "tl";
                case eLightRigDirection.TopRight:
                    return "tr";
                default:
                    return "b";
            }
        }
        #endregion
        #region ePresetColor
        internal static ePresetColor TranslatePresetColor(this string v)
        {
            if (v.Contains("Grey")) v = v.Replace("Grey", "Gray");
            else if (v.Contains("grey")) v = v.Replace("grey", "gray");

            if (v.StartsWith("dk")) v = v.Replace("dk", "Dark");
            else if (v.StartsWith("med")) v = v.Replace("med", "Medium");
            else if (v.StartsWith("lt")) v = v.Replace("lt", "Light");
            return v.ToEnum(ePresetColor.Black);
        }
        internal static string TranslateString(this ePresetColor v)
        {
            var s = v.ToEnumString();
            if (s.StartsWith("dark")) s = s.Replace("dark", "dk");
            else if (s.StartsWith("medium")) s = s.Replace("medium", "med");
            else if (s.StartsWith("light")) s = s.Replace("light", "lt");

            return s;
        }
        #endregion
    }
}
