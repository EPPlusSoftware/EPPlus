/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date (MM/DD/YYYY)              Author                       Change
 *************************************************************************************************
  06/10/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{

    /// <summary>
    /// Manual layout for specifing positions of elements manually.
    /// For easiest use it is recommended to not change the modes of width or height.
    /// Left and Top are used to determine x and y position
    /// Width and Height to define the width and height of the element.
    /// By default all elements originate from their default
    /// Use eLayoutMode.Edge to set origin to the edge of the chart for the relevant element.
    /// </summary>
    public class ExcelManualLayout : XmlHelper
    {
        //string _layoutTargetPath;
        //TODO: Check how this property should be added
        ///// <summary>
        ///// Layout target
        ///// </summary>
        //public eLayoutTarget? LayoutTarget 
        //{
        //    get
        //    {
        //        return GetXmlEnumNull<eLayoutTarget>(_layoutTargetPath);
        //    }
        //    set
        //    {
        //        if ( value == null )
        //        {
        //            DeleteNode(_layoutTargetPath, true);
        //        }
        //        else
        //        {
        //            SetXmlNodeString(_layoutTargetPath, value.ToEnumString());
        //        }
        //    }
        //}

        /// <summary>
        /// Define mode for Left (x) attribute
        /// Edge for origin point left chart edge, Factor for origin point DataLabel position
        /// </summary>
        public eLayoutMode LeftMode
          {
            get
            {
                return GetXmlMode(_path, "x");
            }
            set
            {
                SetXmlMode(_path, "x", value);
            }
        }
        /// <summary>
        /// Define mode for Top (y) attribute
        /// Edge for origin point top chart edge, Factor for origin point DataLabel position
        /// </summary>
        public eLayoutMode TopMode
          {
            get
            {
                return GetXmlMode(_path, "y");
            }
            set
            {
                SetXmlMode(_path, "y", value);
            }
        }
        /// <summary>
        /// Define mode for Width (Right) attribute
        /// Using edge is not recommended.
        /// Edge for Width to be considered the Right of the chart element.
        /// Note: In this case Width will be used for determining Both the element width and its right.
        /// </summary>
        public eLayoutMode WidthMode
          {
            get
            {
                return GetXmlMode(_extLstPath, "w");
            }
            set
            {
                SetXmlMode(_extLstPath, "w", value);
            }
        }
        /// <summary>
        /// Define mode for Height (Bottom) attribute
        /// Using edge is not recommended.
        /// Edge for Height to be considered the bottom of the chart element.
        /// Note: In this case Height will be used for determining Both the element width and its bottom.
        /// </summary>
        public eLayoutMode HeightMode
        {
            get
            {
                return GetXmlMode(_extLstPath, "h");
            }
            set
            {
                SetXmlMode(_extLstPath, "h", value);
            }
        }
        /// <summary>
        /// Define mode for Width (Right) attribute
        /// Using edge is not recommended.
        /// Edge for Width to be considered the Right of the chart element.
        /// Note: In this case Width will be used for determining Both the element width and its right.<para></para>
        /// Legacy variable. if WidthMode property is set this will be overridden.
        /// </summary>
        public eLayoutMode LegacyWidthMode
        {
            get
            {
                return GetXmlMode(_path, "w");
            }
            set
            {
                SetXmlMode(_path, "w", value);
            }
        }
        /// <summary>
        /// Define mode for Height (Bottom) attribute
        /// Using edge is not recommended.
        /// Edge for Height to be considered the bottom of the chart element.<para></para>
        /// Legacy variable. if HeightMode property is set this will be overridden.
        /// </summary>
        public eLayoutMode LegacyHeightMode
        {
            get
            {
                return GetXmlMode(_path, "h");
            }
            set
            {
                SetXmlMode(_path, "h", value);
            }
        }

        /// <summary>
        /// Left offset between 100 to -100% of the chart width. In Excel exceeding these values counts as setting the property to 0.
        /// In Edge mode negative values are not allowed.
        /// </summary>
        public double? Left
        {
            get
            {
                return GetXmlValue(_path, "x");
            }
            set
            {
                SetXmlValue(_path, "x", value);
            }
        }

        /// <summary>
        /// Top offset between 100 to -100% of the chart height. In Excel exceeding these values counts as setting the property to 0.
        /// In Edge mode negative values are not allowed.
        /// </summary>
        public double? Top
        {
            get
            {
                return GetXmlValue(_path, "y");
            }
            set
            {
                SetXmlValue(_path, "y", value);
            }
        }
        /// <summary>
        /// Width offset between 100 to -100% of the chart width. In Excel exceeding these values counts as setting the property to 0.
        /// </summary>
        public double? Width
        {
            get
            {
                return GetXmlValue(_extLstPath, "w");
            }
            set
            {
                if (value != null && WidthMode == eLayoutMode.Edge && value < Left)
                {
                    throw new InvalidOperationException($"Width (Right edge): {value} is less than Left edge {Left}. Cannot invert data label. Right edge cannot pass left edge");
                }
                SetXmlValue(_extLstPath, "w", value);
            }
        }
        /// <summary>
        /// Height offset between 100 to -100% of the chart height. In Excel exceeding these values counts as setting the property to 0.
        /// </summary>
        public double? Height
        {
            get
            {
                return GetXmlValue(_extLstPath, "h");
            }
            set
            {
                if (value != null && HeightMode == eLayoutMode.Edge && value < Top)
                {
                    throw new InvalidOperationException($"Bottom edge (Height) is {value} which is less than Top edge {Top}. Cannot invert element. Right edge cannot pass Left edge");
                }
                SetXmlValue(_extLstPath, "h" ,value);
            }
        }
        /// <summary>
        /// Right offset between 100 to -100% of the chart width. In Excel exceeding these values counts as setting the property to 0.
        /// Legacy variable. if Height property is set this will be overridden.
        /// </summary>
        public double? LegacyWidth
        {
            get
            {
                return GetXmlValue(_path, "w");
            }
            set
            {
                if (value != null && LegacyWidthMode == eLayoutMode.Edge && value < Left)
                {
                    throw new InvalidOperationException($"LegacyWidth (Right edge): {value} is less than Left edge {Left}. Cannot invert data label. Right edge cannot pass left edge");
                }
                SetXmlValue(_path, "w", value);
            }
        }
        /// <summary>
        /// Bottom offset between 100 to -100% of the chart width. In Excel exceeding these values counts as setting the property to 0.
        /// Legacy variable. if Height property is set this will be overridden.
        /// </summary>
        public double? LegacyHeight
        {
            get
            {
                return GetXmlValue(_path,"h");
            }
            set
            {
                if (value != null && LegacyWidthMode == eLayoutMode.Edge && value < Left)
                {
                    throw new InvalidOperationException($"Bottom edge (LegacyHeight) is {value} which is less than Top edge {Top}. Cannot invert data label. Bottom edge cannot pass Top edge");
                }
                SetXmlValue(_path, "h", value);
            }
        }

        private double? GetXmlValue(string path, string name)
        {
            var xmlValue = GetXmlNodeDouble($"{path}/c:{name}/@val");
            return xmlValue == double.NaN ? null : xmlValue * 100;
        }

        private void SetXmlValue(string path, string name, double? value)
        {
            var tempPath = $"{path}/c:{name}/@val";

            if (value == null)
            {
                DeleteNode(tempPath);
            }

            SetXmlNodeString(tempPath, (value.Value * 0.01d).ToString(CultureInfo.InvariantCulture));
        }

        private eLayoutMode GetXmlMode(string path, string name)
        {
            var strValue = GetXmlNodeString($"{path}/c:{name}Mode/@val");
            return strValue == "edge" ? eLayoutMode.Edge : eLayoutMode.Factor;
        }

        private void SetXmlMode(string path, string name, eLayoutMode value)
        {
            var aStr = value.ToEnumString();
            SetXmlNodeString($"{path}/c:{name}Mode/@val", aStr);
        }

        private readonly string _path;
        private readonly string _extLstPath;

        /// <summary>
        /// Manual layout elements
        /// </summary>
        internal ExcelManualLayout(XmlNamespaceManager ns, XmlNode topNode, string path, string extLstPath, string[] schemaNodeOrder = null) : base(ns, topNode) 
        {
            _path = path;
            _extLstPath = extLstPath;
            //_layoutTargetPath = $"{_path}/c:layoutTarget/@val";  Removed for now. See commented out property LayoutTarget above.
            NameSpaceManager.AddNamespace("c15", ExcelPackage.schemaChart2012);
            NameSpaceManager.AddNamespace("c16", ExcelPackage.schemaChart2014);

            AddSchemaNodeOrder(schemaNodeOrder, ["layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h", "extLst"]);
        }

    }
}
