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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
namespace OfficeOpenXml.Drawing.Theme
{

    /// <summary>
    /// The color Scheme for a theme
    /// </summary>
    public class ExcelColorScheme : XmlHelper
    {
        internal ExcelColorScheme(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "dk1","lt1", "dk2", "lt3", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        }
        const string Dk1Path = "a:dk1";
        ExcelDrawingThemeColorManager _dk1 =null;
        /// <summary>
        /// Dark 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Dark1
        {
            get
            {
                if (_dk1 == null)
                {
                    _dk1 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Dk1Path, SchemaNodeOrder);
                }
                return _dk1;
            }
        }

        internal ExcelDrawingThemeColorManager GetColorByEnum(eThemeSchemeColor color)
        {
            switch(color)
            {
                case eThemeSchemeColor.Accent1:
                    return Accent1;
                case eThemeSchemeColor.Accent2:
                    return Accent2;
                case eThemeSchemeColor.Accent3:
                    return Accent3;
                case eThemeSchemeColor.Accent4:
                    return Accent4;
                case eThemeSchemeColor.Accent5:
                    return Accent5;
                case eThemeSchemeColor.Accent6:
                    return Accent6;
                case eThemeSchemeColor.Background1:
                    return Light1;
                case eThemeSchemeColor.Background2:
                    return Light2;
                case eThemeSchemeColor.Text1:
                    return Dark1;
                case eThemeSchemeColor.Text2:
                    return Dark2;
                case eThemeSchemeColor.Hyperlink:
                    return Hyperlink;
                case eThemeSchemeColor.FollowedHyperlink:
                    return FollowedHyperlink;                
            }
            throw(new ArgumentOutOfRangeException($"Type {color} is unhandled."));
        }

        const string Dk2Path = "a:dk2";
        ExcelDrawingThemeColorManager _dk2 = null;
        /// <summary>
        /// Dark 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Dark2
        {
            get
            {
                if (_dk2 == null)
                {
                    _dk2 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Dk2Path, SchemaNodeOrder);
                }
                return _dk2;
            }
        }
        const string lt1Path = "a:lt1";
        ExcelDrawingThemeColorManager _lt1 = null;
        /// <summary>
        /// Light 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Light1
        {
            get
            {
                if (_lt1 == null)
                {
                    _lt1 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, lt1Path, SchemaNodeOrder);
                }
                return _lt1;

            }
        }
        const string lt2Path = "a:lt2";
        ExcelDrawingThemeColorManager _lt2 = null;
        /// <summary>
        /// Light 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Light2
        {
            get
            {
                if (_lt2 == null)
                {
                    _lt2 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, lt2Path, SchemaNodeOrder);
                }
                return _lt2;
            }
        }
        const string Accent1Path = "a:accent1";
        ExcelDrawingThemeColorManager _accent1 = null;
        /// <summary>
        /// Accent 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent1
        {
            get
            {
                if (_accent1 == null)
                {
                    _accent1 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent1Path, SchemaNodeOrder);
                }
                return _accent1;
            }
        }
        const string Accent2Path = "a:accent2";
        ExcelDrawingThemeColorManager _accent2 = null;
        /// <summary>
        /// Accent 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent2
        {
            get
            {
                if (_accent2 == null)
                {
                    _accent2 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent2Path, SchemaNodeOrder);
                }
                return _accent2;
            }
        }
        const string Accent3Path = "a:accent3";
        ExcelDrawingThemeColorManager _accent3 = null;
        /// <summary>
        /// Accent 3 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent3
        {
            get
            {
                if (_accent3 == null)
                {
                    _accent3 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent3Path, SchemaNodeOrder);
                }
                return _accent3;
            }
        }
        const string Accent4Path = "a:accent4";
        ExcelDrawingThemeColorManager _accent4 = null;
        /// <summary>
        /// Accent 4 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent4
        {
            get
            {
                if (_accent4 == null)
                {
                    _accent4 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent4Path, SchemaNodeOrder);
                }
                return _accent4;
            }
        }
        const string Accent5Path = "a:accent5";
        ExcelDrawingThemeColorManager _accent5 = null;
        /// <summary>
        /// Accent 5 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent5
        {
            get
            {
                if (_accent5 == null)
                {
                    _accent5 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent5Path, SchemaNodeOrder);
                }
                return _accent5;
            }
        }
        const string Accent6Path = "a:accent6";
        ExcelDrawingThemeColorManager _accent6 = null;
        /// <summary>
        /// Accent 6 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent6
        {
            get
            {
                if (_accent6 == null)
                {
                    _accent6 = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, Accent6Path, SchemaNodeOrder);
                }
                return _accent6;
            }
        }
        const string HlinkPath = "a:hlink";
        ExcelDrawingThemeColorManager _hlink = null;
        /// <summary>
        /// Hyperlink theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Hyperlink
        {
            get
            {
                if (_hlink == null)
                {
                    _hlink = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, HlinkPath, SchemaNodeOrder);
                }
                return _hlink;
            }
        }

        const string FolHlinkPath = "a:folHlink";
        ExcelDrawingThemeColorManager _folHlink = null;
        /// <summary>
        /// Followed hyperlink theme color
        /// </summary>
        public ExcelDrawingThemeColorManager FollowedHyperlink
        {
            get
            {
                if (_folHlink == null)
                {
                    _folHlink = new ExcelDrawingThemeColorManager(NameSpaceManager, TopNode, FolHlinkPath, SchemaNodeOrder);
                }
                return _folHlink;
            }
        }
    }
}
