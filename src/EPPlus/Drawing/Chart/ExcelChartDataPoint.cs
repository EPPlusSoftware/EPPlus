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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
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
    /// Represents an individual datapoint in a chart
    /// </summary>
    public class ExcelChartDataPoint : XmlHelper, IDisposable, IDrawingStyleBase
    {
        internal const string topNodePath = "c:dPt";
        ExcelChart _chart;
        internal ExcelChartDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
        {
            Init(chart);
            Index = GetXmlNodeInt(indexPath);
        }

        internal ExcelChartDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, int index) : base(ns, topNode)
        {
            Init(chart);
            SetXmlNodeString(indexPath, index.ToString(CultureInfo.InvariantCulture));
            Bubble3D = false;
            Index = index;
        }
        private void Init(ExcelChart chart)
        {
            _chart = chart;
            AddSchemaNodeOrder(new string[] { "idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr", "pictureOptions", "extLst" }, ExcelDrawing._schemaNodeOrderSpPr);
        }
        const string indexPath = "c:idx/@val";
        /// <summary>
        /// The index of the datapoint
        /// </summary>
        public int Index
        {
            get;
            private set;
        }
        /// <summary>
        /// The sizes of the bubbles on the bubble chart
        /// </summary>
        public bool Bubble3D
        {
            get
            {
                return GetXmlNodeBool("c:bubble3D/@val");
            }
            set
            {
                SetXmlNodeString("c:bubble3D/@val", value.GetStringValueForXml());
            }
        }
        /// <summary>
        /// Invert if negative. Default true.
        /// </summary>
        public bool InvertIfNegative
        {
            get
            {
                return GetXmlNodeBool("c:invertIfNegative");
            }
            set
            {
                SetXmlNodeString("c:invertIfNegative", value.GetStringValueForXml());
            }
        }
        ExcelChartMarker _chartMarker = null;
        /// <summary>
        /// A reference to marker properties
        /// </summary>
        public ExcelChartMarker Marker
        {
            get
            {
                if (_chartMarker == null)
                {
                    _chartMarker = new ExcelChartMarker(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _chartMarker;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _line = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_line == null)
                {
                    _line = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
                }
                return _line;
            }
        }
        private ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

        /// <summary>
        /// Returns true if the datapoint has a marker
        /// </summary>
        /// <returns></returns>
        public bool HasMarker()
        {
            return ExistsNode("c:marker");
        }

        /// <summary>
        /// Dispose the object
        /// </summary>
        public void Dispose()
        {
            if (_chart != null) _chart.Dispose();
            _chart = null;
            _line = null;
            if (_fill != null) _fill.Dispose();
            _fill = null;
        }
    }
}
