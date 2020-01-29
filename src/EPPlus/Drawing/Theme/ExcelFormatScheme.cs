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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// The background fill styles, effect styles, fill styles, and line styles which define the style matrix for a theme
    /// </summary>
    public class ExcelFormatScheme : XmlHelper
    {

        private readonly ExcelThemeBase _theme;
        internal ExcelFormatScheme(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            _theme = theme;
        }
        /// <summary>
        /// The name of the format scheme
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        const string fillStylePath = "a:fillStyleLst";
        ExcelThemeFillStyles _fillStyle = null;
        /// <summary>
        ///  Defines the fill styles for the theme
        /// </summary>
        public ExcelThemeFillStyles FillStyle
        {
            get
            {
                if (_fillStyle == null)
                {
                    _fillStyle = new ExcelThemeFillStyles(NameSpaceManager, TopNode.SelectSingleNode(fillStylePath, NameSpaceManager), _theme);
                }
                return _fillStyle;
            }
        }

        const string lineStylePath = "a:lnStyleLst";
        ExcelThemeLineStyles _lineStyle = null;
        /// <summary>
        ///  Defines the line styles for the theme
        /// </summary>
        public ExcelThemeLineStyles BorderStyle
        {
            get
            {
                if (_lineStyle == null)
                {
                    _lineStyle = new ExcelThemeLineStyles(NameSpaceManager, TopNode.SelectSingleNode(lineStylePath, NameSpaceManager));
                }
                return _lineStyle;
            }
        }
        const string effectStylePath = "a:effectStyleLst";
        ExcelThemeEffectStyles _effectStyle = null;
        /// <summary>
        ///  Defines the effect styles for the theme
        /// </summary>
        public ExcelThemeEffectStyles EffectStyle
        {
            get
            {
                if (_effectStyle == null)
                {
                    _effectStyle = new ExcelThemeEffectStyles(NameSpaceManager, TopNode.SelectSingleNode(effectStylePath, NameSpaceManager), _theme);
                }
                return _effectStyle;
            }
        }
        const string backgroundFillStylePath = "a:bgFillStyleLst";
        ExcelThemeFillStyles _backgroundFillStyle = null;
        /// <summary>
        /// Define background fill styles for the theme
        /// </summary>
        public ExcelThemeFillStyles BackgroundFillStyle
        {
            get
            {
                if (_backgroundFillStyle == null)
                {
                    _backgroundFillStyle = new ExcelThemeFillStyles(NameSpaceManager, TopNode.SelectSingleNode(backgroundFillStylePath, NameSpaceManager), _theme);
                }
                return _backgroundFillStyle;
            }
        }
    }
}
