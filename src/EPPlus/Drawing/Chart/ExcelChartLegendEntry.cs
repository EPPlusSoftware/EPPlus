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
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// An individual serie item within the chart legend
    /// </summary>
    public class ExcelChartLegendEntry : XmlHelper, IDrawingStyle
    {
        internal ExcelChartStandard _chart;
        internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStandard chart) : base(nsm, topNode)
        {
            Init(chart);
            Index = GetXmlNodeInt("c:idx/@val");
            HasValue = true;
        }

        internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode legendNode, ExcelChartStandard chart, int serieIndex) : base(nsm)
        {
            Init(chart);
            TopNode = legendNode;
            Index = serieIndex;
        }
        private void Init(ExcelChartStandard chart)
        {
            _chart = chart;
            SchemaNodeOrder = new string[] { "idx", "deleted", "txPr" };
        }
        /// <summary>
        /// The index of the item
        /// </summary>
        public int Index
        {
            get;
            internal set;

        }
        /// <summary>
        /// If the items has been deleted or is visible.
        /// </summary>
        public bool Deleted
        {
            get
            {
                return GetXmlNodeBool("c:delete/@val");
            }
            set
            {
                CreateTopNode();
                HasValue = true;
                SetXmlNodeBool("c:delete/@val", value);
            }
        }
        internal bool HasValue { get; set; }
        private void CreateTopNode()
        {
            if(TopNode.LocalName != "legendEntry")
            {
                var legend = _chart.Legend;
                var preIx = legend.GetPreEntryIndex(Index);
                XmlNode legendEntryNode;
                if (preIx == -1)
                {
                    legendEntryNode = legend.CreateNode("c:legendEntry", false, true);
                }
                else
                {
                    legendEntryNode = _chart.ChartXml.CreateElement("c", "legendEntry", ExcelPackage.schemaChart);
                    var refNode = legend.Entries[preIx].TopNode;
                    refNode.ParentNode.InsertBefore(legendEntryNode, refNode);
                }
                TopNode = legendEntryNode;
                SetXmlNodeInt("c:idx/@val", Index);
            }
        }

        ExcelTextFont _font = null;
        /// <summary>
        /// The Font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    CreateTopNode();
                    _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, $"c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder, InitChartXml);                    
                }
                return _font;
            }
        }
        internal void InitChartXml()
        {
            if (HasValue) return;
            HasValue = true;
            _font.CreateTopNode();
            if (_chart.StyleManager.Style == null) return;
            if (_chart.StyleManager.Style.Legend.HasTextRun)
            {
                var node = (XmlElement)CreateNode("c:txPr/a:p/a:pPr/a:defRPr");
                CopyElement(_chart.StyleManager.Style.Legend.DefaultTextRun.PathElement, node);
            }
            if (_chart.StyleManager.Style.Legend.HasTextBody)
            {
                var node = (XmlElement)CreateNode("c:txPr/a:bodyPr");
                CopyElement(_chart.StyleManager.Style.Legend.DefaultTextBody.PathElement, node);
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, $"c:txPr/a:bodyPr", SchemaNodeOrder);
                }
                return _textBody;
            }
        }

        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                return null;
            }
        }
        /// <summary>
        /// Access to effects styling properties
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Access to fill styling properties.
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Access to 3D properties.
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                return null;
            }
        }


        internal void Save()
        {
            if(Deleted==true)
            {
                DeleteNode("c:txPr");
            }
            else
            {
                if (ExistsNode("c:txPr"))
                {
                    DeleteNode("c:delete");
                }
                else
                {
                    TopNode.ParentNode.RemoveChild(TopNode);
                }
            }
        }

        public void CreatespPr()
        {
            
        }
    }
}