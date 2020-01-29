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
using System.Drawing;
using System.Xml;
using System;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Manages colors in a theme 
    /// </summary>
    public class ExcelDrawingThemeColorManager
    {
        /// <summary>
        /// Namespace manager
        /// </summary>
        internal protected XmlNamespaceManager _nameSpaceManager;
        /// <summary>
        /// The top node
        /// </summary>
        internal protected XmlNode _topNode;
        /// <summary>
        /// The node of the supplied path
        /// </summary>
        internal protected XmlNode _pathNode = null;
        /// <summary>
        /// The node of the color object
        /// </summary>
        internal protected XmlNode _colorNode = null;
        /// <summary>
        /// Init method
        /// </summary>
        internal protected Action _initMethod;
        /// <summary>
        /// The x-path
        /// </summary>
        internal protected string _path;
        /// <summary>
        /// Order of the elements according to the xml schema
        /// </summary>
        internal protected string[] _schemaNodeOrder;
        internal ExcelDrawingThemeColorManager(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, Action initMethod = null)
        {
            _nameSpaceManager = nameSpaceManager;
            _topNode = topNode;
            _path = path;
            _initMethod = initMethod;
            _pathNode = GetPathNode();
            _schemaNodeOrder = schemaNodeOrder;
            if (_pathNode == null) return;

            if (IsTopNodeColorNode(_topNode))
            {
                _colorNode = _pathNode;
            }
            else
            {
                _colorNode = _pathNode.FirstChild;
            }

            if (_colorNode == null)
            {
                return;
            }

            switch (_colorNode.LocalName)
            {
                case "sysClr":
                    ColorType = eDrawingColorType.System;
                    SystemColor = new ExcelDrawingSystemColor(_nameSpaceManager, _pathNode.FirstChild);
                    break;
                case "scrgbClr":
                    ColorType = eDrawingColorType.RgbPercentage;
                    RgbPercentageColor = new ExcelDrawingRgbPercentageColor(_nameSpaceManager, _pathNode.FirstChild);
                    break;
                case "hslClr":
                    ColorType = eDrawingColorType.Hsl;
                    HslColor = new ExcelDrawingHslColor(_nameSpaceManager, _pathNode.FirstChild);
                    break;
                case "prstClr":
                    ColorType = eDrawingColorType.Preset;
                    PresetColor = new ExcelDrawingPresetColor(_nameSpaceManager, _pathNode.FirstChild);
                    break;
                case "srgbClr":
                    ColorType = eDrawingColorType.Rgb;
                    RgbColor = new ExcelDrawingRgbColor(_nameSpaceManager, _pathNode.FirstChild);
                    break;
                default:
                    ColorType = eDrawingColorType.None;
                    break;
            }
        }

        private bool IsTopNodeColorNode(XmlNode topNode)
        {
            return topNode.LocalName.EndsWith("Clr");                
        }

        /// <summary>
        /// The type of color.
        /// Each type has it's own property and set-method.       
        /// <see cref="SetRgbColor(Color, bool)"/>
        /// <see cref="SetRgbPercentageColor(double, double, double)"/>
        /// <see cref="SetHslColor(double, double, double)" />
        /// <see cref="SetPresetColor(Color)"/>
        /// <see cref="SetPresetColor(ePresetColor)"/>
        /// <see cref="SetSystemColor(eSystemColor)"/>
        /// <see cref="ExcelDrawingColorManager.SetSchemeColor(eSchemeColor)"/>
        /// </summary>
        public eDrawingColorType ColorType { get; internal protected set; } = eDrawingColorType.None;
        internal void SetXml(XmlNamespaceManager nameSpaceManager, XmlNode node)
        {
            
        }
        ExcelColorTransformCollection _transforms = null;
        /// <summary>
        /// Color transformations
        /// </summary>
        public ExcelColorTransformCollection Transforms
        {
            get
            {
                if (ColorType == eDrawingColorType.None) return null;
                if (_transforms == null)
                {
                    _transforms = new ExcelColorTransformCollection(_nameSpaceManager, _colorNode);
                }
                return _transforms;
            }
        }
        /// <summary>
        /// A rgb color.
        /// This property has a value when Type is set to Rgb
        /// </summary>
        public ExcelDrawingRgbColor RgbColor { get; private set; }
        /// <summary>
        /// A rgb precentage color.
        /// This property has a value when Type is set to RgbPercentage
        /// </summary>
        public ExcelDrawingRgbPercentageColor RgbPercentageColor { get; private set; }
        /// <summary>
        /// A hsl color.
        /// This property has a value when Type is set to Hsl
        /// </summary>
        public ExcelDrawingHslColor HslColor { get; private set; }
        /// <summary>
        /// A preset color.
        /// This property has a value when Type is set to Preset
        /// </summary>
        public ExcelDrawingPresetColor PresetColor { get; private set; }
        /// <summary>
        /// A system color.
        /// This property has a value when Type is set to System
        /// </summary>
        public ExcelDrawingSystemColor SystemColor { get; private set; }
        /// <summary>
        /// Sets a rgb color.
        /// </summary>
        /// <param name="color">The color</param>
        /// <param name="setAlpha">Apply the alpha part of the Color to the <see cref="Transforms"/> collection</param>
        public void SetRgbColor(Color color, bool setAlpha=false)
        {
            ColorType = eDrawingColorType.Rgb;
            ResetColors(ExcelDrawingRgbColor.NodeName);

            if(setAlpha && color.A != 0xFF)
            {
                Transforms.RemoveOfType(eColorTransformType.Alpha);
                Transforms.AddAlpha((double)color.A);
            }
            RgbColor = new ExcelDrawingRgbColor(_nameSpaceManager, _colorNode) { Color = color };
        }
        /// <summary>
        /// Sets a rgb precentage color
        /// </summary>
        /// <param name="redPercentage">Red percentage</param>
        /// <param name="greenPercentage">Green percentage</param>
        /// <param name="bluePercentage">Bluepercentage</param>
        public void SetRgbPercentageColor(double redPercentage, double greenPercentage, double bluePercentage)
        {
            ColorType = eDrawingColorType.RgbPercentage;
            ResetColors(ExcelDrawingRgbPercentageColor.NodeName);
            RgbPercentageColor = new ExcelDrawingRgbPercentageColor(_nameSpaceManager, _colorNode) { RedPercentage = redPercentage, GreenPercentage = greenPercentage, BluePercentage = bluePercentage };
        }
        /// <summary>
        /// Sets a hsl color
        /// </summary>
        /// <param name="hue">The hue angle. From 0-360</param>
        /// <param name="saturation">The saturation percentage. From 0-100</param>
        /// <param name="luminance">The luminance percentage. From 0-100</param>
        public void SetHslColor(double hue, double saturation, double luminance)
        {
            ColorType = eDrawingColorType.Hsl;
            ResetColors(ExcelDrawingHslColor.NodeName);
            HslColor = new ExcelDrawingHslColor(_nameSpaceManager, _colorNode) { Hue = hue, Saturation= saturation, Luminance = luminance };
        }
        /// <summary>
        /// Sets a preset color.
        /// Must be a named color. Can't be color.Empty.
        /// </summary>
        /// <param name="color">Color</param>
        public void SetPresetColor(Color color)
        {
            ColorType = eDrawingColorType.Preset;
            ResetColors(ExcelDrawingPresetColor.NodeName);
            PresetColor = new ExcelDrawingPresetColor(_nameSpaceManager, _colorNode) { Color = ExcelDrawingPresetColor.GetPresetColor(color) };
        }
        /// <summary>
        /// Sets a preset color.
        /// </summary>
        /// <param name="presetColor">The color</param>
        public void SetPresetColor(ePresetColor presetColor)
        {
            ColorType = eDrawingColorType.Preset;
            ResetColors(ExcelDrawingPresetColor.NodeName);
            PresetColor = new ExcelDrawingPresetColor(_nameSpaceManager, _colorNode) { Color = presetColor };
        }
        /// <summary>
        /// Sets a system color
        /// </summary>
        /// <param name="systemColor">The colors</param>
        public void SetSystemColor(eSystemColor systemColor)
        {
            ColorType = eDrawingColorType.System;
            ResetColors(ExcelDrawingSystemColor.NodeName);
            SystemColor = new ExcelDrawingSystemColor(_nameSpaceManager, _colorNode) { Color = systemColor };
        }
        /// <summary>
        /// Reset the color objects
        /// </summary>
        /// <param name="newNodeName">The new color node name</param>
        internal protected virtual void ResetColors(string newNodeName)
        {
            if(_colorNode==null)
            {
                var xml = XmlHelperFactory.Create(_nameSpaceManager, _topNode);
                xml.SchemaNodeOrder = _schemaNodeOrder;
                var colorPath = string.IsNullOrEmpty(_path) ? newNodeName : _path + "/" + newNodeName;
                _colorNode = xml.CreateNode(colorPath);
                _initMethod?.Invoke();
            }
            if (_colorNode.Name == newNodeName)
            {
                return;
            }
            else
            {
                _transforms = null;
                ChangeType(newNodeName);
            }

            RgbColor = null;
            RgbPercentageColor = null;
            HslColor = null;
            PresetColor = null;
            SystemColor = null;
        }

        private void ChangeType(string type)
        {
            if(_topNode==_colorNode)
            {
                var xh = XmlHelperFactory.Create(_nameSpaceManager, _topNode);
                xh.ReplaceElement(_colorNode, type);
            }
            else
            {
                var p = _colorNode.ParentNode;
                p.InnerXml = $"<{type} />";
                _colorNode = p.FirstChild;
            }
        }
        private XmlNode GetPathNode()
        {
            if (_pathNode == null)
            {
                if (string.IsNullOrEmpty(_path))
                {
                    _pathNode = _topNode;
                }
                else
                {
                    _pathNode = _topNode.SelectSingleNode(_path, _nameSpaceManager);
                }

            }
            return _pathNode;
        }

    }
}