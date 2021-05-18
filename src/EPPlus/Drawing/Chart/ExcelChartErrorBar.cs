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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public class ExcelChartErrorBars : XmlHelper, IDrawingStyleBase
    {
        readonly ExcelChartSerieWithErrorBars _chartSerie;
        internal ExcelChartErrorBars(ExcelChartSerieWithErrorBars chartSerie) :
            this(chartSerie, chartSerie.TopNode)
        {
        }
        internal ExcelChartErrorBars(ExcelChartSerieWithErrorBars chartSerie, XmlNode topNode) :
            base(chartSerie.NameSpaceManager, topNode)
        {
            _chartSerie = chartSerie;
            AddSchemaNodeOrder(new string[]{ "errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "val", "spPr" }, ExcelDrawing._schemaNodeOrderSpPr);
            if (TopNode.LocalName != "errBars")
            {
                TopNode = chartSerie.CreateNode("c:errBars", false, true);
            }
        }
        string _directionPath = "c:errDir/@val";
        /// <summary>
        /// The directions for the error bars. For scatter-, bubble- and area charts this property can't be changed. Please use the ErrorBars property for Y direction and ErrorBarsX for the X direction.
        /// </summary>
        public eErrorBarDirection Direction
        {
            get
            {
                ValidateNotDeleted();
                return GetXmlNodeString(_directionPath).ToEnum(eErrorBarDirection.Y);
            }
            set
            {
                ValidateNotDeleted();
                if(_chartSerie._chart.IsTypeBubble() || _chartSerie._chart.IsTypeScatter() || _chartSerie._chart.IsTypeArea())
                {
                    if(value!=Direction)
                    {
                        throw new InvalidOperationException("Can't change direction for this chart type. Please use ErrorBars or ErrorBarsX property to determin the direction.");
                    }
                    return;
                }
                SetDirection(value);
            }
        }

        internal void SetDirection(eErrorBarDirection value)
        {
            SetXmlNodeString(_directionPath, value.ToEnumString());
        }

        string _barTypePath = "c:errBarType/@val";
        /// <summary>
        /// The ways to draw an error bar
        /// </summary>
        public eErrorBarType BarType
        {
            get
            {
                ValidateNotDeleted();
                return GetXmlNodeString(_barTypePath).ToEnum(eErrorBarType.Both);
            }
            set
            {
                ValidateNotDeleted();
                SetXmlNodeString(_barTypePath, value.ToEnumString());
            }
        }
        string _valueTypePath = "c:errValType/@val";
        /// <summary>
        /// The ways to determine the length of the error bars
        /// </summary>
        public eErrorValueType ValueType
        {
            get
            {
                ValidateNotDeleted();
                return GetXmlNodeString(_valueTypePath).TranslateErrorValueType();
            }
            set
            {
                ValidateNotDeleted();
                SetXmlNodeString(_valueTypePath, value.ToEnumString());
            }
        }        
        string _noEndCapPath = "c:noEndCap/@val";
        /// <summary>
        /// If true, no end cap is drawn on the error bars 
        /// </summary>
        public bool NoEndCap
        {
            get
            {
                ValidateNotDeleted();
                return GetXmlNodeBool(_noEndCapPath, true);
            }
            set
            {
                ValidateNotDeleted();
                SetXmlNodeBool(_noEndCapPath, value, true);
            }
        }

        string _valuePath = "c:val/@val";
        /// <summary>
        /// The value which used to determine the length of the error bars when <c>ValueType</c> is FixedValue
        /// </summary>
        public double? Value
        {
            get
            {
                ValidateNotDeleted();
                return GetXmlNodeDoubleNull(_valuePath);
            }
            set
            {
                ValidateNotDeleted();
                if (value == null)
                {
                    DeleteNode(_valuePath, true);
                }
                else
                {
                    SetXmlNodeString(_valuePath, value.Value.ToString("R15", CultureInfo.InvariantCulture));
                }
            }
        }
        string _plusNodePath = "c:plus";
        ExcelChartNumericSource _plus=null;
        /// <summary>
        /// Numeric Source for plus errorbars when <c>ValueType</c> is set to Custom
        /// </summary>
        public ExcelChartNumericSource Plus
        {
            get
            {
                ValidateNotDeleted();
                if (_plus==null)
                {
                    _plus = new ExcelChartNumericSource(NameSpaceManager, TopNode, _plusNodePath, SchemaNodeOrder);
                }
                return _plus;
            }
        }
        string _minusNodePath = "c:minus";
        ExcelChartNumericSource _minus = null;
        /// <summary>
        /// Numeric Source for minus errorbars when <c>ValueType</c> is set to Custom
        /// </summary>
        public ExcelChartNumericSource Minus
        {
            get
            {
                ValidateNotDeleted();
                if (_minus == null)
                {
                    _minus = new ExcelChartNumericSource(NameSpaceManager, TopNode, _minusNodePath, SchemaNodeOrder);
                }
                return _minus;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Fill style
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                ValidateNotDeleted();
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chartSerie._chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border style
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                ValidateNotDeleted();
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chartSerie._chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
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
                ValidateNotDeleted();
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chartSerie._chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
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
                ValidateNotDeleted();
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        private void ValidateNotDeleted()
        {
            if(TopNode==null)
            {
                throw new InvalidOperationException("The error bar has been deleted.");
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

        /// <summary>
        /// Remove the error bars
        /// </summary>
        public void Remove()
        {
            DeleteNode(".");
            if(_chartSerie.ErrorBars==this)
            {
                _chartSerie.ErrorBars = null;
            }
            if(_chartSerie is ExcelChartSerieWithHorizontalErrorBars errorBarsSerie)
            {
                if (errorBarsSerie.ErrorBarsX == this)
                {
                    errorBarsSerie.ErrorBarsX = null;
                }
            }
        }
    }
}
