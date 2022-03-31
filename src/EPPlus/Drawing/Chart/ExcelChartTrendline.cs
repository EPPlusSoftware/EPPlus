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
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.ThreeD;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A trendline object
    /// </summary>
    public class ExcelChartTrendline : XmlHelper, IDrawingStyleBase
    {
        ExcelChartStandardSerie _serie;
        internal ExcelChartTrendline(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelChartStandardSerie serie) :
            base(namespaceManager,topNode)

        {
            _serie = serie;
            AddSchemaNodeOrder(new string[] { "name", "spPr", "trendlineType","order","period", "forward","backward","intercept", "dispRSqr", "dispEq", "trendlineLbl" }, ExcelDrawing._schemaNodeOrderSpPr);
        }
        const string TRENDLINEPATH = "c:trendlineType/@val";
        /// <summary>
        /// Type of Trendline
        /// </summary>
        public eTrendLine Type
        {
           get
           {
               switch (GetXmlNodeString(TRENDLINEPATH).ToLower(CultureInfo.InvariantCulture))
               {
                   case "exp":
                       return eTrendLine.Exponential;
                   case "log":
                        return eTrendLine.Logarithmic;
                   case "poly":
                       return eTrendLine.Polynomial;
                   case "movingavg":
                       return eTrendLine.MovingAvgerage;
                   case "power":
                       return eTrendLine.Power;
                   default:
                       return eTrendLine.Linear;
               }
           }
           set
           {
                switch (value)
                {
                    case eTrendLine.Exponential:
                        SetXmlNodeString(TRENDLINEPATH, "exp");
                        break;
                    case eTrendLine.Logarithmic:
                        SetXmlNodeString(TRENDLINEPATH, "log");
                        break;
                    case eTrendLine.Polynomial:
                        SetXmlNodeString(TRENDLINEPATH, "poly");
                        Order = 2;
                        break;
                    case eTrendLine.MovingAvgerage:
                        SetXmlNodeString(TRENDLINEPATH, "movingAvg");
                        Period = 2;
                        break;
                    case eTrendLine.Power:
                        SetXmlNodeString(TRENDLINEPATH, "power");
                        break;
                    default: 
                        SetXmlNodeString(TRENDLINEPATH, "linear");
                        break;
                }
           }
        }
        const string NAMEPATH = "c:name";
        /// <summary>
        /// Name in the legend
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAMEPATH);
            }
            set
            {
                SetXmlNodeString(NAMEPATH, value, true);
            }
        }
        const string ORDERPATH = "c:order/@val";
        /// <summary>
        /// Order for polynominal trendlines
        /// </summary>
        public decimal Order
        {
            get
            {
                return GetXmlNodeDecimal(ORDERPATH);
            }
            set
            {
                if (Type == eTrendLine.MovingAvgerage)
                {
                    throw (new ArgumentException("Can't set period for trendline type MovingAvgerage"));
                }
                DeleteAllNode(PERIODPATH);
                SetXmlNodeString(ORDERPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string PERIODPATH = "c:period/@val";
        /// <summary>
        /// Period for monthly average trendlines
        /// </summary>
        public decimal Period
        {
            get
            {
                return GetXmlNodeDecimal(PERIODPATH);
            }
            set
            {
                if (Type == eTrendLine.Polynomial)
                {
                    throw (new ArgumentException("Can't set period for trendline type Polynomial"));
                }
                DeleteAllNode(ORDERPATH);
                SetXmlNodeString(PERIODPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string FORWARDPATH = "c:forward/@val";
        /// <summary>
        /// Forcast forward periods
        /// </summary>
        public decimal Forward
        {
            get
            {
                return GetXmlNodeDecimal(FORWARDPATH);
            }
            set
            {
                SetXmlNodeString(FORWARDPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string BACKWARDPATH = "c:backward/@val";
        /// <summary>
        /// Forcast backwards periods
        /// </summary>
        public decimal Backward
        {
            get
            {
                return GetXmlNodeDecimal(BACKWARDPATH);
            }
            set
            {
                SetXmlNodeString(BACKWARDPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string INTERCEPTPATH = "c:intercept/@val";
        /// <summary>
        /// The point where the trendline crosses the vertical axis
        /// </summary>
        public decimal Intercept
        {
            get
            {
                return GetXmlNodeDecimal(INTERCEPTPATH);
            }
            set
            {
                SetXmlNodeString(INTERCEPTPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string DISPLAYRSQUAREDVALUEPATH = "c:dispRSqr/@val";
        /// <summary>
        /// If to display the R-squared value for a trendline
        /// </summary>
        public bool DisplayRSquaredValue
        {
            get
            {
                return GetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH, true);
            }
            set
            {
                SetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH, value, true);
            }
        }
        const string DISPLAYEQUATIONPATH = "c:dispEq/@val";
        /// <summary>
        /// If to display the trendline equation on the chart
        /// </summary>
        public bool DisplayEquation
        {
            get
            {
                return GetXmlNodeBool(DISPLAYEQUATIONPATH, true);
            }
            set
            {
                SetXmlNodeBool(DISPLAYEQUATIONPATH, value, true);
            }
        }
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
                    _fill = new ExcelDrawingFill(_serie._chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
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
                    _border = new ExcelDrawingBorder(_serie._chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_serie._chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
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

        ExcelChartTrendlineLabel _label =null;
        /// <summary>
        /// Trendline labels
        /// </summary>
        public ExcelChartTrendlineLabel Label
        {
            get
            {
                if(_label==null)
                {
                    _label = new ExcelChartTrendlineLabel(NameSpaceManager, TopNode, _serie);
                }
                return _label;
            }
        }

        /// <summary>
        /// Return true if the trendline has labels.
        /// </summary>
        public bool HasLbl
        {
            get
            {
                return ExistsNode("c:trendlineLbl") ||
                       (Type != eTrendLine.MovingAvgerage && (DisplayRSquaredValue == true || DisplayEquation == true));
            }
        }
    }
}
