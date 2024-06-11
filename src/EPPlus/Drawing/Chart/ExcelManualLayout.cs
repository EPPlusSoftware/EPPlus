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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{

    /// <summary>
    /// Manual layout for specifing positions of label elements manually
    /// </summary>
    public class ExcelManualLayout : XmlHelper
    {
        eLayoutTarget layoutTarget;

        eLayoutMode xMode;
        eLayoutMode yMode;
        eLayoutMode wMode;
        eLayoutMode hMode;

        /// <summary>
        /// x position as a percentual value based on the parent object and the layout mode.
        /// Keep in mind that if your label is centered -50 will be the maximum left value and + 50 the maximum right
        /// </summary>
        public double x
        {
            get
            {
                return GetXmlNodeInt($"{_path}/c:x/@val");
            }
            set
            {
                SetXmlNodeString($"{_path}/c:x/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// y position as a percentual value based on the parent object and the layout mode.
        /// </summary>
        public double y
        {
            get
            {
                return GetXmlNodeInt($"{_path}/c:y/@val");
            }
            set
            {
                SetXmlNodeString($"{_path}/c:y/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }        
        /// <summary>
        /// width as a percentual value based on the parent object and the layout mode.
        /// specifies right if wMode is Edge
        /// </summary>
        public double w
        {
            get
            {
                return GetXmlNodeInt($"{_path}/c:w/@val");
            }
            set
            {
                SetXmlNodeString($"{_path}/c:w/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }        
        /// <summary>
        /// height as a percentual value based on the parent object and the layout mode.
        /// specifies bottom if hMode is edge 
        /// </summary>
        public double h
        {
            get
            {
                return GetXmlNodeInt($"{_path}/c:h/@val");
            }
            set
            {
                SetXmlNodeString($"{_path}/c:h/@val", (value * 0.01d).ToString(CultureInfo.InvariantCulture));
            }
        }

        private readonly string _path;

        /// <summary>
        /// Manual layout elements
        /// </summary>
        public ExcelManualLayout(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder = null) : base(ns, topNode) 
        {
            _path = path;
            AddSchemaNodeOrder(schemaNodeOrder, ["layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h", "extLst"]);
        }

    }
}
