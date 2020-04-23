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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extentions;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public class ExcelChartErrorBars : XmlHelper, IDrawingStyleBase
    {
        ExcelChart _chart;
        string[] _topSchemaNodeOrder;
        internal ExcelChartErrorBars(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode node, string[] schemaNodeOrder) :
            base(nameSpaceManager, node)
        {
            _chart = chart;
            _topSchemaNodeOrder = schemaNodeOrder;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "val", "spPr" }, new int[] { 0, schemaNodeOrder.Length });
            AddSchemaNodeOrder(SchemaNodeOrder, ExcelDrawing._schemaNodeOrderSpPr);
        }
        string _directionPath = "c:errBars/c:errDir/@val";
        /// <summary>
        /// The directions for the error bars
        /// </summary>
        public eErrorBarDirection Direction
        {
            get
            {
                return GetXmlNodeString(_directionPath).ToEnum(eErrorBarDirection.Y);
            }
            set
            {
                SetXmlNodeString(_directionPath, value.ToEnumString());
            }
        }
        string _barTypePath = "c:errBars/c:errBarType/@val";
        /// <summary>
        /// The ways to draw an error bar
        /// </summary>
        public eErrorBarType BarType
        {
            get
            {
                return GetXmlNodeString(_barTypePath).ToEnum(eErrorBarType.Both);
            }
            set
            {
                SetXmlNodeString(_barTypePath, value.ToEnumString());
            }
        }
        string _valueTypePath = "c:errBars/c:errValType/@val";
        /// <summary>
        /// The ways to determine the length of the error bars
        /// </summary>
        public eErrorValueType ValueType
        {
            get
            {
                return GetXmlNodeString(_valueTypePath).TranslateErrorValueType();
            }
            set
            {
                SetXmlNodeString(_valueTypePath, value.ToEnumString());
            }
        }        
        string _noEndCapPath = "c:errBars/c:noEndCap/@val";
        /// <summary>
        /// If true, no end cap is drawn on the error bars 
        /// </summary>
        public bool NoEndCap
        {
            get
            {
                return GetXmlNodeBool(_noEndCapPath, true);
            }
            set
            {
                SetXmlNodeBool(_noEndCapPath, value, true);
            }
        }

        string _valuePath = "c:errBars/c:val/@val";
        /// <summary>
        /// The value which used to determine the length of the error bars when <c>ValueType</c> is FixedValue
        /// </summary>
        public double? Value
        {
            get
            {
                return GetXmlNodeDoubleNull(_valuePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_valuePath);
                }
                else
                {
                    SetXmlNodeString(_valuePath, value.Value.ToString("R15", CultureInfo.InvariantCulture));
                }
            }
        }
        string _plusNodePath = "c:errBars/c:plus";
        ExcelChartNumericSource _plus=null;
        /// <summary>
        /// Numeric Source for plus errorbars when <c>ValueType</c> is set to Custom
        /// </summary>
        public ExcelChartNumericSource Plus
        {
            get
            {
                if(_plus==null)
                {
                    _plus = new ExcelChartNumericSource(NameSpaceManager, TopNode, _plusNodePath, SchemaNodeOrder);
                }
                return _plus;
            }
        }
        string _minusNodePath = "c:errBars/c:minus";
        ExcelChartNumericSource _minus = null;
        /// <summary>
        /// Numeric Source for minus errorbars when <c>ValueType</c> is set to Custom
        /// </summary>
        public ExcelChartNumericSource Minus
        {
            get
            {
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
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "c:errBars/c:spPr", SchemaNodeOrder);
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
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "c:errBars/c:spPr/a:ln", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "c:errBars/c:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:errBars/c:spPr", SchemaNodeOrder);
                }
                return _threeD;
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
            DeleteNode("c:errBars");
        }
    }
}
