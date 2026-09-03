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
using OfficeOpenXml.Drawing.Shape.Style;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// A reference from a chart style to the theme collection
    /// </summary>
    public class ExcelChartStyleReference : ShapeStyleReference
    {
        string _path;
        internal ExcelChartStyleReference(XmlNamespaceManager nsm, XmlNode topNode, string path) : base(nsm, topNode, path)
        {
            _path = path;            
        }

        ExcelChartStyleColorManager _color = null;
        /// <summary>
        /// The color to be used for the reference. 
        /// This will replace any the StyleClr node in the chart style xml.
        /// </summary>
        public ExcelChartStyleColorManager Color
        {
            get
            {
                if(_color==null)
                {
                    _color = new ExcelChartStyleColorManager(NameSpaceManager, TopNode, _path, SchemaNodeOrder);
                }

                return _color;
            }
        }
    }
}