/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Individual color settings for a region map charts series colors
    /// </summary>
    public class ExcelChartExValueColor : XmlHelper
    {
        string _prefix;
        string _positionPath;
        internal ExcelChartExValueColor(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string prefix) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _prefix = prefix;
            _positionPath = $"cx:valueColorPositions/cx:{prefix}Position";
        }

        ExcelDrawingColorManager _color = null;
        /// <summary>
        /// The color
        /// </summary>
        public ExcelDrawingColorManager Color
        {
            get
            {
                if(_color==null)
                {
                    _color = new ExcelDrawingColorManager(NameSpaceManager, TopNode, $"cx:valueColors/cx:{_prefix}Color", SchemaNodeOrder);
                }
                return _color;
            }
        }
        /// <summary>
        /// The color variation type.
        /// </summary>
        public eColorValuePositionType ValueType
        {
            get
            {
                if(ExistsNode($"{_positionPath}/cx:number"))
                {
                    return eColorValuePositionType.Number;
                }
                else if(ExistsNode($"{_positionPath}/cx:percent"))
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
                    ClearChildren(_positionPath);
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

        /// <summary>
        /// The color variation value.
        /// </summary>
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
                    return GetXmlNodeDoubleNull($"{_positionPath}/cx:percent/@val")??0;
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
                    if (value < 0 || value > 100)
                    {
                        throw new InvalidOperationException("PositionValue out of range. Percantage range is from 0 to 100");
                    }
                    SetXmlNodeDouble($"{_positionPath}/cx:percent/@val", value);
                }
            }
        }
    }
}