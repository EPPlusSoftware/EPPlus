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
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A charts plot area
    /// </summary>
    public class ExcelChartPlotArea :  XmlHelper, IDrawingStyleBase
    {
        ExcelChart _firstChart;
        ExcelChart _topChart;
        string _nsPrefix;
        internal ExcelChartPlotArea(XmlNamespaceManager ns, XmlNode node, ExcelChart firstChart, string nsPrefix, ExcelChart topChart=null)
           : base(ns,node)
       {
            _nsPrefix = nsPrefix;
            if(firstChart._isChartEx)
            {
                AddSchemaNodeOrder(new string[] { "plotAreaRegion", "plotSurface", "series", "axis","spPr" },
                    ExcelDrawing._schemaNodeOrderSpPr);
            }
            else
            {
                AddSchemaNodeOrder(new string[] { "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart", "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart", "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart", "valAx", "catAx", "dateAx", "serAx", "dTable", "spPr" },
                    ExcelDrawing._schemaNodeOrderSpPr);
            }

            _firstChart = firstChart;
            _topChart = topChart ?? firstChart;
            if (TopNode.SelectSingleNode("c:dTable", NameSpaceManager) != null)
            {
                DataTable = new ExcelChartDataTable(firstChart,NameSpaceManager, TopNode);
            }
        }

        ExcelChartCollection _chartTypes;
        /// <summary>
        /// If a chart contains multiple chart types (e.g lineChart,BarChart), they end up here.
        /// </summary>
        public ExcelChartCollection ChartTypes  
        {
            get
            {
                if (_chartTypes == null)
                {
                    _chartTypes = new ExcelChartCollection(_topChart);
                    _chartTypes.Add(_firstChart);
                    if (_topChart!=_firstChart)
                    {
                        _chartTypes.Add(_topChart);
                    }
                }
                return _chartTypes;
            }
        }
        #region Data table
        /// <summary>
        /// Creates a data table in the plotarea
        /// The datatable can also be accessed via the DataTable propery
        /// <see cref="DataTable"/>
        /// </summary>
        public virtual ExcelChartDataTable CreateDataTable()
        {
            if(DataTable!=null)
            {
                throw (new InvalidOperationException("Data table already exists"));
            }

            DataTable = new ExcelChartDataTable(_firstChart, NameSpaceManager, TopNode);
            return DataTable;
        }
        /// <summary>
        /// Remove the data table if it's created in the plotarea
        /// </summary>
        public virtual void RemoveDataTable()
        {
            DeleteAllNode("c:dTable");
            DataTable = null;
        }
        /// <summary>
        /// The data table object.
        /// Use the CreateDataTable method to create a datatable if it does not exist.
        /// <see cref="CreateDataTable"/>
        /// <see cref="RemoveDataTable"/>
        /// </summary>
        public ExcelChartDataTable DataTable { get; private set; } = null;
        #endregion
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_firstChart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_firstChart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:ln", SchemaNodeOrder);
                }   
                return _border;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_firstChart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }
    }
}
