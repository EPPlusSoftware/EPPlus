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
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// An effect style for a theme
    /// </summary>
    public class ExcelThemeEffectStyle : XmlHelper
    {
        string _path;
        string[] _schemaNodeOrder;
        private readonly ExcelThemeBase _theme;
        internal ExcelThemeEffectStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            if (!string.IsNullOrEmpty(path)) path += "/";
            _path = path;
            _schemaNodeOrder = schemaNodeOrder;
            _theme = theme;
        }
        ExcelDrawingEffectStyle _effects = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if(_effects==null)
                {
                    _effects = new ExcelDrawingEffectStyle(_theme, NameSpaceManager, TopNode, _path + "a:effectLst", _schemaNodeOrder);
                }
                return _effects;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D settings
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, _path, _schemaNodeOrder);
                }
                return _threeD;
            }
        }
    }
}
