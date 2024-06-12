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
    /// Manual layout for specifing positions of label elements manually.
    /// </summary>
    public class ExcelManualLayout : XmlHelper
    {
        eLayoutTarget layoutTarget;

        public eLayoutMode LeftMode
          {
            get
            {
                var strValue = GetXmlNodeString($"{_path}/c:xMode/@val");
                return strValue == "edge" ? eLayoutMode.Edge : eLayoutMode.Factor;
            }
            set
            {
                SetXmlNodeString($"{_path}/c:xMode/@val", value.ToEnumString());
            }
        }
        public eLayoutMode TopMode
          {
            get
            {
                var strValue = GetXmlNodeString($"{_path}/c:yMode/@val");
                return strValue == "edge" ? eLayoutMode.Edge : eLayoutMode.Factor;
            }
            set
            {
                SetXmlNodeString($"{_path}/c:yMode/@val", value.ToEnumString());
            }
        }
        public eLayoutMode WidthMode
          {
            get
            {
                var strValue = GetXmlNodeString($"{_path}/c:wMode/@val");
                return strValue == "edge" ? eLayoutMode.Edge : eLayoutMode.Factor;
            }
            set
            {
                SetXmlNodeString($"{_path}/c:wMode/@val", value.ToEnumString());
            }
        }
        public eLayoutMode HeightMode
        {
            get
            {
                var strValue = GetXmlNodeString($"{_path}/c:hMode/@val");
                return strValue == "edge" ? eLayoutMode.Edge : eLayoutMode.Factor;
            }
            set
            {
                var aStr = value.ToEnumString();
                SetXmlNodeString($"{_path}/c:hMode/@val", aStr);
            }
        }

        /// <summary>
        /// Left offset between 100 to -100%. In Excel exceeding these values counts as setting the property to 0.
        /// 
        /// </summary>
        public double Left
        {
            get
            {
                return GetXmlNodeDouble($"{_path}/c:x/@val") * 100;
            }
            set
            {
                SetXmlNodeString($"{_path}/c:x/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// -100 to 100% offset from top of the chart or relative to the default position if hMode is Factor
        /// Going above chart limits counts as setting the value to 0 visually.
        /// </summary>
        public double Top
        {
            get
            {
                return GetXmlNodeDouble($"{_path}/c:y/@val") * 100;
            }
            set
            {
                SetXmlNodeString($"{_path}/c:y/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }        
        /// <summary>
        /// Width of the textbox around the label. As a positive percentual value of the chart.
        /// 100 = 100% of the chart width.
        /// width as a percentual value based on the parent object and the layout mode.
        /// specifies right if wMode is Edge
        /// </summary>
        public double Width
        {
            get
            {
                return GetXmlNodeDouble($"{_path}/c:w/@val") * 100;
            }
            set
            {
                if(WidthMode == eLayoutMode.Edge && value < Left)
                {
                    throw new InvalidOperationException($"Width (Right edge): {value} is less than Left edge {Left}. Cannot invert data label. Right edge cannot pass left edge");
                }
                SetXmlNodeString($"{_path}/c:w/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }
        /// <summary>
        /// Height of the textbox around the label. As a positive percentual value of the chart.
        /// 100 = 100% of the chart Height.
        /// height as a percentual value based on the parent object and the layout mode.
        /// specifies bottom if hMode is edge 
        /// </summary>
        public double Height
        {
            get
            {
                return GetXmlNodeDouble($"{_path}/c:h/@val") * 100;
            }
            set
            {
                if (HeightMode == eLayoutMode.Edge && value < Top)
                {
                    throw new InvalidOperationException($"Bottom edge (Height) is {value} which is less than Top edge {Top}. Cannot invert data label. Bottom edge cannot pass Top edge");
                }
                SetXmlNodeString($"{_path}/c:h/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }

        private readonly string _path;

        /// <summary>
        /// Manual layout elements
        /// </summary>
        internal ExcelManualLayout(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder = null) : base(ns, topNode) 
        {
            _path = path;
            AddSchemaNodeOrder(schemaNodeOrder, ["layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h", "extLst"]);
        }

    }
}
