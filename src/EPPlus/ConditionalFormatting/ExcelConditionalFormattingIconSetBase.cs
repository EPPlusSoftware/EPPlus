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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// IconSet base class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class ExcelConditionalFormattingIconSetBase<T> : 
        ExcelConditionalFormattingRule,
        IExcelConditionalFormattingThreeIconSet<T>
        where T : struct, Enum
    {
        private string _uid = null;

        internal override string Uid { 
            get 
            {
                if(_uid == null)
                {
                    return NewId();
                }

                return _uid;
            } 
            set
            {
                _uid = value;
            }
        }

        internal static string NewId()
        {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        internal ExcelConditionalFormattingIconSetBase(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
            :base(type, 
                 address, 
                 priority, 
                 worksheet) 
        {
            double symbolCount;

            if (type == eExcelConditionalFormattingRuleType.ThreeIconSet)
            {
                symbolCount = 3;
            }
            else if(type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                symbolCount = 4;
            }
            else
            {
                symbolCount = 5;
            }

            Icon1 = CreateIcon(0, type);
            Icon2 = CreateIcon(Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(Math.Round(100D * (2D / symbolCount), 0), type);
        }

        protected ExcelConditionalFormattingIconDataBarValue CreateIcon(double value, eExcelConditionalFormattingRuleType type)
        {
            var icon = new ExcelConditionalFormattingIconDataBarValue
                (
                eExcelConditionalFormattingValueObjectType.Percent,
                type
                );

            icon.Value = value;

            return icon;
        }

        internal ExcelConditionalFormattingIconSetBase(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet,
          bool stopIfTrue,
          XmlReader xr)
            :base (type, address, priority, worksheet)
        {
            StopIfTrue = stopIfTrue;

            ShowValue = xr.GetAttribute("showValue") != "0";
            IconSetPercent = xr.GetAttribute("percent") != "0";
            Reverse = xr.GetAttribute("reverse") == "0";

            var set = xr.GetAttribute("iconSet").Substring(1);

            Type = type;
            IconSet = set.ToEnum<T>().Value;

            double symbolCount;

            if (type == eExcelConditionalFormattingRuleType.ThreeIconSet)
            {
                symbolCount = 3;
            }
            else if (type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                symbolCount = 4;
            }
            else
            {
                symbolCount = 5;
            }

            Icon1 = CreateIcon(0, type);
            Icon2 = CreateIcon(Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(Math.Round(100D * (2D / symbolCount), 0), type);

            xr.Read();
            Icon1.Type = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            Icon1.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
            Icon2.Type = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            Icon2.Value = double.Parse(xr.GetAttribute("val"));

            if (!string.IsNullOrEmpty(xr.GetAttribute("gte")))
            {
                Icon2.GreaterThanOrEqualTo = int.Parse(xr.GetAttribute("gte")) != 0;
            }

            xr.Read();
            Icon3.Type = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            Icon3.Value = double.Parse(xr.GetAttribute("val"));
            if (!string.IsNullOrEmpty(xr.GetAttribute("gte")))
            {
                Icon3.GreaterThanOrEqualTo = int.Parse(xr.GetAttribute("gte")) != 0;
            }

            xr.Read();
        }

        internal ExcelConditionalFormattingIconSetBase(ExcelConditionalFormattingIconSetBase<T> copy) : base(copy)
        {
            StopIfTrue = copy.StopIfTrue;
            ShowValue = copy.ShowValue;
            IconSetPercent = copy.IconSetPercent;
            Reverse = copy.Reverse;

            Type = copy.Type;
            IconSet = copy.IconSet;

            Icon1 = copy.Icon1;
            Icon2 = copy.Icon2;
            Icon3 = copy.Icon3;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingIconSetBase<T>(this);
        }

        /// <summary>
        /// Settings for icon 1 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon1
        {
            get;
            internal set;
        }

        /// <summary>
        /// Settings for icon 2 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon2
        {
            get;
            internal set;
        }
        /// <summary>
        /// Settings for icon 2 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon3
        {
            get;
            internal set;
        }

        /// <summary>
        /// Reverse the order of the icons
        /// Default false
        /// </summary>
        public bool Reverse
        {
            get;
            set;
        }

        /// <summary>
        /// If its percent
        /// default true
        /// </summary>
        public bool IconSetPercent
        {
            get;
            set;
        }

        /// <summary>
        /// If the cell values are visible
        /// default true
        /// </summary>
        public bool ShowValue
        {
            get;
            set;
        }

        internal override bool IsExtLst
        {
            get
            {
                if (GetIconSetString() == "3Stars" ||
                    GetIconSetString() == "3Triangles" ||
                    GetIconSetString() == "5Boxes")
                {
                    return true;
                }

                if(ExcelAddressBase.RefersToOtherWorksheet(Icon1.Formula, _ws.Name))
                {
                    return true;
                }

                if (ExcelAddressBase.RefersToOtherWorksheet(Icon2.Formula, _ws.Name))
                {
                    return true;
                }

                if (ExcelAddressBase.RefersToOtherWorksheet(Icon3.Formula, _ws.Name))
                {
                    return true;
                }

                return false;
            }
        }

        public T IconSet
        {
            get;
            set;
        }

        internal string GetIconSetString()
        {
            return GetIconSetString(IconSet);
        }

        internal string GetIconSetString(T value)
        {
            if (Type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "4Arrows";
                    case "ArrowsGray":
                        return "4ArrowsGray";
                    case "Rating":
                        return "4Rating";
                    case "RedToBlack":
                        return "4RedToBlack";
                    case "TrafficLights":
                        return "4TrafficLights";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
            else if (Type == eExcelConditionalFormattingRuleType.FiveIconSet)
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "5Arrows";
                    case "ArrowsGray":
                        return "5ArrowsGray";
                    case "Quarters":
                        return "5Quarters";
                    case "Rating":
                        return "5Rating";
                    case "Boxes":
                        return "5Boxes";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
            else
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "3Arrows";
                    case "ArrowsGray":
                        return "3ArrowsGray";
                    case "Flags":
                        return "3Flags";
                    case "Signs":
                        return "3Signs";
                    case "Symbols":
                        return "3Symbols";
                    case "Symbols2":
                        return "3Symbols2";
                    case "TrafficLights1":
                        return "3TrafficLights1";
                    case "TrafficLights2":
                        return "3TrafficLights2";
                    case "Stars":
                        return "3Stars";
                    case "Triangles":
                        return "3Triangles";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
        }
    }
}