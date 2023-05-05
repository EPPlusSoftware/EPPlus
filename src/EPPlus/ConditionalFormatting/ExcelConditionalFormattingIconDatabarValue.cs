using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// 18.3.1.11 cfvo (Conditional Format Value Object)
    /// Describes the values of the interpolation points in a gradient scale.
    /// </summary>
    public class ExcelConditionalFormattingIconDataBarValue
    {
        private eExcelConditionalFormattingRuleType _ruleType;
        ExcelConditionalFormattingRule _rule;

        internal bool HasValueOrFormula
        {
            get
            {
                if (Type != eExcelConditionalFormattingValueObjectType.Min
                    && Type != eExcelConditionalFormattingValueObjectType.Max)
                {
                    return true;
                }

                return false;
            }
        }

        //eExcelConditionalFormattingValueObjectType _valueType;

        internal int minLength = 0;
        internal int maxLength = 100;

        internal ExcelConditionalFormattingIconDataBarValue(
            eExcelConditionalFormattingValueObjectType valueType,
            eExcelConditionalFormattingRuleType ruleType)
        {
            RuleType = ruleType;
            Type = valueType;
        }

        /// <summary>
        /// If not custom is null. If user assigns to it holds icon value.
        /// </summary>
        public eExcelconditionalFormattingCustomIcon? CustomIcon { get; set; } = null;

        readonly Dictionary<int, string> _iconStringSetDictionary = new Dictionary<int, string>
            {
             { 0,  "3Arrows" },
             { 1,  "3ArrowsGray" },
             { 2,  "3Flags" },
             { 3,  "3TrafficLights1" } ,
             { 4,  "3TrafficLights2" },
             { 5,  "3Signs" },
             { 6,  "3Symbols" },
             { 7,  "3Symbols2" },
             { 8,  "3Stars" },
             { 9,  "3Triangles" },
             { 10, "4Arrows" },
             { 11, "4ArrowsGray" },
             { 12, "4RedToBlack" },
             { 13, "4Rating" },
             { 14, "4TrafficLights" },
             { 15, "5Rating" },
             { 16, "5Quarters" },
             { 17, "5Boxes" },
             { 18, "NoIcons"},
            };

        internal void SetCustomIconStringAndId(string set, int id)
        {
            int myKey = _iconStringSetDictionary.FirstOrDefault(x => x.Value == set).Key << 4;
            myKey += id;
            CustomIcon = (eExcelconditionalFormattingCustomIcon)myKey;
        }

        internal virtual string GetCustomIconStringValue()
        {
            if (CustomIcon != null)
            {
                int customIconId = (int)CustomIcon;

                var iconSetId = customIconId >> 4;

                return _iconStringSetDictionary[iconSetId];
            }

            throw new NotImplementedException($"Cannot get custom icon {CustomIcon} of {this} ");
        }

        internal int GetCustomIconIndex()
        {
            if (CustomIcon != null)
            {
                return (int)CustomIcon & 0xf;
            }

            return -1;
        }

        /// <summary>
        /// Rule type
        /// </summary>
        internal eExcelConditionalFormattingRuleType RuleType
        {
            get { return _ruleType; }
            set { _ruleType = value; }
        }

        eExcelConditionalFormattingValueObjectType _type;

        /// <summary>
        /// Value type
        /// </summary>
        public eExcelConditionalFormattingValueObjectType Type
        {
            get
            {
                return _type;
            }
            set
            {
                if ((_ruleType == eExcelConditionalFormattingRuleType.ThreeIconSet || _ruleType == eExcelConditionalFormattingRuleType.FourIconSet || _ruleType == eExcelConditionalFormattingRuleType.FiveIconSet) &&
                    (value == eExcelConditionalFormattingValueObjectType.Min || value == eExcelConditionalFormattingValueObjectType.Max))
                {
                    throw new ArgumentException("Value type can't be Min or Max for iconSets");
                }

                _type = value;
            }
        }

        /// <summary>
        /// Greater Than Or Equal To
        /// Set to false to only apply an icon when greaterThan
        /// </summary>
        public bool GreaterThanOrEqualTo { get; set; } = true;

        private double? _value = double.NaN;

        /// <summary>
        /// The value
        /// </summary>
        public double Value
        {
            get
            {
                if (Type == eExcelConditionalFormattingValueObjectType.Num
                    || Type == eExcelConditionalFormattingValueObjectType.Percent
                    || Type == eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    return (double)_value;
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                _value = null;

                // Only some types use the @val attribute
                if (Type == eExcelConditionalFormattingValueObjectType.Num
                    || Type == eExcelConditionalFormattingValueObjectType.Percent
                    || Type == eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    _value = value;
                }
                else
                {
                    throw new InvalidOperationException("Value can only be changed if Type is Num, Percent or Percentile." +
                        $"Current Type is \"{Type}\"");
                }
            }
        }

        string _formula = "";

        /// <summary>
        /// The Formula of the Object Value (uses the same attribute as the Value)
        /// </summary>
        public string Formula
        {
            get
            {
                // Return empty if the Object Value type is not Formula
                if (Type != eExcelConditionalFormattingValueObjectType.Formula)
                {
                    return string.Empty;
                }

                // Excel stores the formula in the @val attribute
                return _formula;
            }
            set
            {
                // Only store the formula if the Object Value type is Formula
                if (Type == eExcelConditionalFormattingValueObjectType.Formula)
                {
                    _formula = value;
                }
                else
                {
                    throw new InvalidOperationException("Cannot store formula in a non-formula type");
                }
            }
        }
    }
}
