/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelChartEx : ExcelChartBase
    {
        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, bool isPivot, ExcelGroupShape parent) : 
            base(drawings, node, GetChartType(node, drawings.NameSpaceManager), isPivot, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {
            _isChartEx = true;
        }

        private static eChartType GetChartType(XmlNode node, XmlNamespaceManager nsm)
        {
            var layoutId = node.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[1]/@layoutId", nsm);
            if (layoutId == null) throw new InvalidOperationException($"No series in chart"); 
            switch(layoutId.Value)
            {
                case "clusteredColumn":
                    return eChartType.Histogram;
                case "paretoLine":
                    return eChartType.Pareto;
                case "boxWhisker":
                    return eChartType.Boxwhisker;
                case "funnel":
                    return eChartType.Funnel;
                case "regionMap":
                    return eChartType.RegionMap;
                case "sunburst":
                    return eChartType.Sunburst;
                case "treemap":
                    return eChartType.Treemap;
                case "waterfall":
                    return eChartType.Waterfall;
                default:
                    throw new InvalidOperationException($"Unsupported layoutId in ChartEx Xml: {layoutId}");
            }          
        }

        internal override void AddAxis()
        {
            var l = new List<ExcelChartAxis>();
            foreach (XmlNode axNode in _chartXmlHelper.GetNodes("cx:plotArea/cx:axis"))
            {
                l.Add(new ExcelChartExAxis(this, NameSpaceManager, axNode));
            }
        }

        internal ExcelChartEx(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
            base(drawings, drawingsNode, type, null, PivotTableSource, chartXml, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {
            _isChartEx = true;
        }


        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent=null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {
            _isChartEx = true;
            ChartType = GetChartType(chartNode, drawings.NameSpaceManager);
            Series.Init(this, drawings.NameSpaceManager, chartNode, false, Series._list);
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public override ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr/a:ln", _chartXmlHelper.SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public override ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public override ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr/a:effectLst", _chartXmlHelper.SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public override ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public override ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:txPr/a:p/a:pPr/a:defRPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _font;
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public override ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:txPr/a:bodyPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _textBody;
            }
        }
        /// <summary>
        /// Chart series
        /// </summary>
        public new ExcelChartSeries<ExcelChartExSerie> Series { get; } = new ExcelChartSeries<ExcelChartExSerie>();

        public override bool VaryColors
        {
            get 
            { 
                return false; 
            }
            set
            {
                throw new InvalidOperationException("VaryColors do not apply to Extended charts");
            }
        }
    }
}
