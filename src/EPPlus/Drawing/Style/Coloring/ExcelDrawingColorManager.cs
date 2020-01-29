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
using System.Xml;
using System;
using System.Linq;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Handles colors for drawings
    /// </summary>
    public class ExcelDrawingColorManager : ExcelDrawingThemeColorManager
    {
        internal ExcelDrawingColorManager(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, Action initMethod = null) : 
            base(nameSpaceManager, topNode, path, schemaNodeOrder, initMethod)
        {
            if (_pathNode == null || _colorNode==null)  return;
         
            switch (_colorNode.LocalName)
            {
                case "schemeClr":
                    ColorType = eDrawingColorType.Scheme;
                    SchemeColor = new ExcelDrawingSchemeColor(_nameSpaceManager, _colorNode);
                    break;
            }
        }
        /// <summary>
        /// If <c>type</c> is set to SchemeColor, then this property contains the scheme color
        /// </summary>
        public ExcelDrawingSchemeColor SchemeColor { get; private set; }
        /// <summary>
        /// Set the color to a scheme color
        /// </summary>
        /// <param name="schemeColor">The scheme color</param>
        public void SetSchemeColor(eSchemeColor schemeColor)
        {
            ColorType = eDrawingColorType.Scheme;
            ResetColors(ExcelDrawingSchemeColor.NodeName);
            SchemeColor = new ExcelDrawingSchemeColor(_nameSpaceManager, _colorNode) { Color=schemeColor };
        }
        /// <summary>
        /// Reset the colors on the object
        /// </summary>
        /// <param name="newNodeName">The new color new name</param>
        internal new protected void ResetColors(string newNodeName) 
        {
            base.ResetColors(newNodeName);
            SchemeColor = null;
        }

        internal void ApplyNewColor(ExcelDrawingColorManager newColor, ExcelColorTransformCollection variation=null)
        {
            ColorType = newColor.ColorType;
            switch (newColor.ColorType)
            {
                case eDrawingColorType.Rgb:
                    SetRgbColor(newColor.RgbColor.Color);
                    break;
                case eDrawingColorType.RgbPercentage:
                    SetRgbPercentageColor(newColor.RgbPercentageColor.RedPercentage, newColor.RgbPercentageColor.GreenPercentage, newColor.RgbPercentageColor.BluePercentage);
                    break;
                case eDrawingColorType.Hsl:
                    SetHslColor(newColor.HslColor.Hue, newColor.HslColor.Saturation, newColor.HslColor.Luminance);
                    break;
                case eDrawingColorType.Preset:
                    SetPresetColor(newColor.PresetColor.Color);
                    break;
                case eDrawingColorType.System:
                    SetSystemColor(newColor.SystemColor.Color);
                    break;
                case eDrawingColorType.Scheme:
                    SetSchemeColor(newColor.SchemeColor.Color);
                    break;
            }
            //Variations should be added first, so temporary store the transforms and add the again
            var trans = Transforms.ToList();
            Transforms.Clear();
            if (variation != null)
            {
                ApplyNewTransform(variation);
            }
            ApplyNewTransform(trans);
            ApplyNewTransform(newColor.Transforms);
        }

        private void ApplyNewTransform(IEnumerable<IColorTransformItem> transforms)
        {
            foreach (var t in transforms)
            {
                switch(t.Type)
                {
                    case eColorTransformType.Alpha:
                        Transforms.AddAlpha(t.Value);
                        break;
                    case eColorTransformType.AlphaMod:
                        Transforms.AddAlphaModulation(t.Value);
                        break;
                    case eColorTransformType.AlphaOff:
                        Transforms.AddAlphaOffset(t.Value);
                        break;
                    case eColorTransformType.Blue:
                        Transforms.AddBlue(t.Value);
                        break;
                    case eColorTransformType.BlueMod:
                        Transforms.AddBlueModulation(t.Value);
                        break;
                    case eColorTransformType.BlueOff:
                        Transforms.AddBlueOffset(t.Value);
                        break;
                    case eColorTransformType.Comp:
                        Transforms.AddComplement();
                        break;
                    case eColorTransformType.Gamma:
                        Transforms.AddGamma();
                        break;
                    case eColorTransformType.Gray:
                        Transforms.AddGray();
                        break;
                    case eColorTransformType.Green:
                        Transforms.AddGreen(t.Value);
                        break;
                    case eColorTransformType.GreenMod:
                        Transforms.AddGreenModulation(t.Value);
                        break;
                    case eColorTransformType.GreenOff:
                        Transforms.AddGreenOffset(t.Value);
                        break;
                    case eColorTransformType.Hue:
                        Transforms.AddHue(t.Value);
                        break;
                    case eColorTransformType.HueMod:
                        Transforms.AddHueModulation(t.Value);
                        break;
                    case eColorTransformType.HueOff:
                        Transforms.AddHueOffset(t.Value);
                        break;
                    case eColorTransformType.Inv:
                        Transforms.AddInverse();
                        break;
                    case eColorTransformType.InvGamma:
                        Transforms.AddGamma();
                        break;
                    case eColorTransformType.Lum:
                        Transforms.AddLuminance(t.Value);
                        break;
                    case eColorTransformType.LumMod:
                        Transforms.AddLuminanceModulation(t.Value);
                        break;
                    case eColorTransformType.LumOff:
                        Transforms.AddLuminanceOffset(t.Value);
                        break;
                    case eColorTransformType.Red:
                        Transforms.AddRed(t.Value);
                        break;
                    case eColorTransformType.RedMod:
                        Transforms.AddRedModulation(t.Value);
                        break;
                    case eColorTransformType.RedOff:
                        Transforms.AddRedOffset(t.Value);
                        break;
                    case eColorTransformType.Sat:
                        Transforms.AddSaturation(t.Value);
                        break;
                    case eColorTransformType.SatMod:
                        Transforms.AddSaturationModulation(t.Value);
                        break;
                    case eColorTransformType.SatOff:
                        Transforms.AddSaturationOffset(t.Value);
                        break;
                    case eColorTransformType.Shade:
                        Transforms.AddShade(t.Value);
                        break;
                    case eColorTransformType.Tint:
                        Transforms.AddTint(t.Value);
                        break;
                }
            }
        }
    }
}