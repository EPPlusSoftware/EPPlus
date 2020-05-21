/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExValueColor : XmlHelper
    {
        string _prefix;
        string _positionPath;
        internal ExcelChartExValueColor(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string prefix) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _prefix = prefix;
            _positionPath = $"cx:valueColorPosition/cx{prefix}Position";
        }

        ExcelDrawingColorManager _color = null;
        public ExcelDrawingColorManager Color
        {
            get
            {
                if(_color==null)
                {
                    _color = new ExcelDrawingColorManager(NameSpaceManager, TopNode, $"cx:valueColors/{_prefix}Color", SchemaNodeOrder);
                }
                return _color;
            }
        }
        public eColorValuePositionType ValueType
        {
            get
            {
                if(ExistNode($"{_positionPath}/cx:number"))
                {
                    return eColorValuePositionType.Number;
                }
                else if(ExistNode($"{_positionPath}/cx:percent"))
                {
                    return eColorValuePositionType.Percent;
                }
                else
                {
                    return eColorValuePositionType.Extreme;
                }
            }
            set
            {
                if(ValueType!=value)
                {
                    DeleteNode(_positionPath);
                    switch(value)
                    {
                        case eColorValuePositionType.Extreme:
                            CreateNode($"{_positionPath}/cx:extremeValue");
                            break;
                        case eColorValuePositionType.Percent:
                            SetXmlNodeString($"{_positionPath}/cx:percent/@val", "0");
                            break;
                        default:
                            SetXmlNodeString($"{_positionPath}/cx:number/@val", "0");
                            break;
                    }
                }
            }
        }

        public double PositionValue
        {
            get
            {
                var t = ValueType;
                if (t==eColorValuePositionType.Extreme)
                {
                    return 0;
                }
                else if(ValueType==eColorValuePositionType.Number)
                {
                    return GetXmlNodeDouble($"{_positionPath}/cx:number/@val");
                }
                else
                {
                    return GetXmlNodePercentage($"{_positionPath}/cx:number/@val")??0;
                }
            }
            set
            {
                var t = ValueType;
                if (t==eColorValuePositionType.Extreme)
                {
                    throw (new InvalidOperationException("Can't set PositionValue when ValueType is Extreme"));
                }
                else if (t==eColorValuePositionType.Number)
                {
                    SetXmlNodeString($"{_positionPath}/cx:number/@val", value.ToString(CultureInfo.InvariantCulture));
                }
                else if (t == eColorValuePositionType.Percent)
                {
                    SetXmlNodePercentage($"{_positionPath}/cx:percent/@val", value, false);
                }
            }
        }
    }
    public class ExcelChartExValueColors : XmlHelper
    {
        private ExcelRegionMapChartSerie _series;

        internal ExcelChartExValueColors(ExcelRegionMapChartSerie series, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder) : base (nameSpaceManager, topNode)
        {
            _series = series;
            SchemaNodeOrder = schemaNodeOrder;
        }

        ExcelChartExValueColor _minColor = null;
        public ExcelChartExValueColor MinColor 
        {
            get
            {
                if(_minColor==null)
                {
                    _minColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "min");
                }
                return _minColor;
            }
        }
        ExcelChartExValueColor _midColor = null;
        public ExcelChartExValueColor MidColor
        {
            get
            {
                if (_midColor == null)
                {
                    _midColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "mid");
                }
                return _midColor;
            }
        }
        ExcelChartExValueColor _maxColor = null;
        public ExcelChartExValueColor MaxColor
        {
            get
            {
                if (_maxColor == null)
                {
                    _maxColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "max");
                }
                return _maxColor;
            }
        }
    }
}