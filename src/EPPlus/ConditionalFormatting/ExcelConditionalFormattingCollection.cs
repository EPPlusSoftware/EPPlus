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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingCollection : IEnumerable<IExcelConditionalFormattingRule>
    {
        List<ExcelConditionalFormattingRule> _rules = new List<ExcelConditionalFormattingRule>();
        ExcelWorksheet _ws;
        int LastPriority = 1;
        internal Dictionary<string, ExcelConditionalFormattingRule> _extLstDict = new Dictionary<string, ExcelConditionalFormattingRule>();
        List<ExcelConditionalFormattingRule> _dataBarStorage = new List<ExcelConditionalFormattingRule>();


        internal ExcelConditionalFormattingCollection(ExcelWorksheet ws)
        {
            _ws = ws;
            _rules = new List<ExcelConditionalFormattingRule>();
        }

        /// <summary>
        /// Read conditionalFormatting info from extLst in xml via xr reader
        /// </summary>
        public void ReadExtConditionalFormattings(XmlReader xr)
        {
            while (xr.Read())
            {
                //Localname should always be 'conditionalFormatting' if another node or 'conditionalFormattings' if finished
                if (xr.LocalName != "conditionalFormatting")
                {
                    xr.Read(); //Read beyond the end element
                    break;
                }

                if (xr.NodeType == XmlNodeType.Element)
                {
                    //ConditionalFormatting->cfRule
                    xr.Read();

                    var addresslessCFs = new List<ExcelConditionalFormattingRule>();  
                    do
                    {
                        string id = xr.GetAttribute("id");

                        if (string.IsNullOrEmpty(id))
                        {
                            throw new InvalidOperationException("XML invalid. cfRule without Id found");
                        }

                        if (xr.GetAttribute("type") == "dataBar")
                        {
                            //cfRule->Type
                            xr.Read();

                            var dataBar = (ExcelConditionalFormattingDataBar)_extLstDict[id];
                            dataBar.LowValue.minLength = int.Parse(xr.GetAttribute("minLength"));
                            dataBar.HighValue.maxLength = int.Parse(xr.GetAttribute("maxLength"));

                            //CfRule -> cfvo
                            xr.Read();

                            string typeString1 = RemoveAuto(xr.GetAttribute("type"));

                            dataBar.LowValue.Type = typeString1.ToEnum<eExcelConditionalFormattingValueObjectType>().Value;

                            xr.Read();

                            if (dataBar.LowValue.HasValueOrFormula)
                            {
                                xr.Read();
                                if (dataBar.LowValue.Type == eExcelConditionalFormattingValueObjectType.Formula)
                                {
                                    dataBar.LowValue.Formula = xr.ReadContentAsString();
                                }
                                else
                                {
                                    dataBar.LowValue.Value = double.Parse(xr.ReadContentAsString());
                                }
                                xr.Read();
                                xr.Read();
                            }

                            string typeString2 = RemoveAuto(xr.GetAttribute("type"));

                            dataBar.HighValue.Type = typeString2.ToEnum<eExcelConditionalFormattingValueObjectType>().Value;

                            xr.Read();

                            if (dataBar.HighValue.HasValueOrFormula)
                            {
                                xr.Read();
                                if (dataBar.HighValue.Type == eExcelConditionalFormattingValueObjectType.Formula)
                                {
                                    dataBar.HighValue.Formula = xr.ReadContentAsString();
                                }
                                else
                                {
                                    dataBar.HighValue.Value = double.Parse(xr.ReadContentAsString());
                                }
                                xr.Read();
                                xr.Read();
                            }

                            if (xr.LocalName == "fillColor")
                            {
                                ReadCT_Color(xr, dataBar.FillColor);
                            }

                            if (xr.LocalName == "borderColor")
                            {
                                ReadCT_Color(xr, dataBar.BorderColor);
                            }

                            if (xr.LocalName == "negativeFillColor")
                            {
                                ReadCT_Color(xr, dataBar.NegativeFillColor);
                            }

                            if (xr.LocalName == "negativeBorderColor")
                            {
                                ReadCT_Color(xr, dataBar.NegativeBorderColor);
                            }

                            if (xr.LocalName == "axisColor")
                            {
                                ReadCT_Color(xr, dataBar.AxisColor);
                            }

                            // /DataBar-> /cfRule -> xm:sqref -> textValue
                            xr.Read();
                            xr.Read();
                            if (xr.LocalName != "cfRule")
                            {
                                xr.Read();
                                //If we need to handle ext adress it can be read here with xr.ReadContentAsString();
                                // textValue -> /xm:sqref -> /conditionalFormatting
                                xr.Read();
                                xr.Read();
                            }

                            _rules.Add(dataBar);
                        }
                        else if (xr.GetAttribute("type") == "iconSet")
                        {
                            //cfRule->Type
                            xr.Read();

                            string iconSet = xr.GetAttribute("iconSet");

                            int numIcons = int.Parse(iconSet[0].ToString());

                            //iconSet -> cfvo
                            xr.Read();

                            var types = new List<string>();
                            var values = new List<double>();

                            for (int i = 0; i < numIcons; i++)
                            {
                                types.Add(xr.GetAttribute("type"));

                                xr.Read();
                                xr.Read();

                                values.Add(double.Parse(xr.Value));

                                xr.Read();
                                xr.Read();
                                xr.Read();
                            }

                            List<string> customIconTypes = null;
                            List<int> customIconIds = null;

                            if (xr.LocalName == "cfIcon")
                            {
                                customIconTypes = new List<string>();
                                customIconIds = new List<int>();

                                for (int i = 0; i < numIcons; i++)
                                {
                                    customIconTypes.Add(xr.GetAttribute("iconSet"));
                                    customIconIds.Add(int.Parse(xr.GetAttribute("iconId")));
                                    xr.Read();
                                }
                            }

                            //var dataBar = (ExcelConditionalFormattingDataBar)_extLstDict[id];

                            //iconSet->cfRule->sqref
                            string address = null;
                            xr.Read();
                            xr.Read();

                            if (xr.LocalName != "cfRule")
                            {
                                xr.Read();

                                address = xr.ReadContentAsString();

                                xr.Read();

                                if (customIconTypes == null)
                                {
                                    xr.Read();
                                }
                            }

                            ExcelAddress iconAddress = null;
                            if (address != null)
                            {
                                iconAddress = new ExcelAddress(address);
                            }

                            ExcelConditionalFormattingRule rule = null;

                            switch (iconSet[0])
                            {
                                case '3':

                                    IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> threeIconSet;

                                    if (iconAddress != null)
                                    {
                                        threeIconSet = AddThreeIconSet(
                                         iconAddress,
                                         iconSet.Substring(1).ToEnum<eExcelconditionalFormatting3IconsSetType>().Value);
                                    }
                                    else
                                    {
                                        threeIconSet = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)
                                            AddRule(
                                                    eExcelConditionalFormattingRuleType.ThreeIconSet,
                                                    iconAddress, true);

                                        threeIconSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting3IconsSetType>().Value;

                                        UpdateExtDict();
                                    }

                                    ApplyIconSetExtValues(
                                        new ExcelConditionalFormattingIconDataBarValue[]
                                        { threeIconSet.Icon1, threeIconSet.Icon2, threeIconSet.Icon3 },
                                        types, values, customIconTypes, customIconIds);

                                    rule = (ExcelConditionalFormattingRule)threeIconSet;

                                    break;

                                case '4':

                                    IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> fourSet;

                                    if (iconAddress != null)
                                    {
                                        fourSet = AddFourIconSet(
                                        iconAddress,
                                        iconSet.Substring(1).ToEnum<eExcelconditionalFormatting4IconsSetType>().Value);
                                    }
                                    else
                                    {
                                        fourSet = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)
                                         AddRule(
                                                 eExcelConditionalFormattingRuleType.FourIconSet,
                                                 iconAddress, true);

                                        fourSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting4IconsSetType>().Value;

                                        UpdateExtDict();
                                    }

                                    ApplyIconSetExtValues(
                                    new ExcelConditionalFormattingIconDataBarValue[]
                                    { fourSet.Icon1, fourSet.Icon2, fourSet.Icon3, fourSet.Icon4 },
                                    types, values, customIconTypes, customIconIds);

                                    rule = (ExcelConditionalFormattingRule)fourSet;

                                    break;

                                case '5':
                                    IExcelConditionalFormattingFiveIconSet fiveSet;

                                    if (iconAddress != null)
                                    {
                                        fiveSet = AddFiveIconSet(
                                         iconAddress,
                                         iconSet.Substring(1).ToEnum<eExcelconditionalFormatting5IconsSetType>().Value);
                                    }
                                    else
                                    {
                                        fiveSet = (IExcelConditionalFormattingFiveIconSet)
                                         AddRule(
                                                 eExcelConditionalFormattingRuleType.FiveIconSet,
                                                 iconAddress, true);

                                        fiveSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting5IconsSetType>().Value;

                                        UpdateExtDict();
                                    }

                                    ApplyIconSetExtValues(
                                     new ExcelConditionalFormattingIconDataBarValue[]
                                     { fiveSet.Icon1, fiveSet.Icon2, fiveSet.Icon3, fiveSet.Icon4 , fiveSet.Icon5 },
                                     types, values, customIconTypes, customIconIds);

                                    rule = (ExcelConditionalFormattingRule)fiveSet;
                                    break;
                            }


                            if (iconAddress == null && rule != null)
                            {
                                addresslessCFs.Add(rule);
                            }
                        }
                        else
                        {
                            var cf = ExcelConditionalFormattingRuleFactory.Create(null, _ws, xr);
                            _rules.Add(cf);
                            _extLstDict.Add(cf.Uid, cf);

                            if (cf.Address == null)
                            {
                                addresslessCFs.Add(cf);
                            }
                        }
                    } while (xr.LocalName == "cfRule");

                    foreach (var cf in addresslessCFs)
                    {
                        cf.Address = _rules.LastOrDefault().Address;
                    }
                }
            }
        }

        void ReadCT_Color(XmlReader xr, ExcelDxfColor color)
        {

            if (!string.IsNullOrEmpty(xr.GetAttribute("theme")))
            {
                color.Theme = (eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("rgb")))
            {
                color.Color = GetColorFromExcelRgb(xr.GetAttribute("rgb"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("auto")))
            {
                color.Auto = xr.GetAttribute("auto") == "1" ? true : false;
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("index")))
            {
                color.Index = int.Parse(xr.GetAttribute("index"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
            {
                color.Tint = double.Parse(xr.GetAttribute("tint"));
            }

            xr.Read();
        }

        ExcelConditionalFormattingIconDataBarValue[] CreateBaseIconArr(eExcelConditionalFormattingRuleType type)
        {
            int nrOfIcons;
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    nrOfIcons = 3;
                    break;
                case eExcelConditionalFormattingRuleType.FourIconSet:
                    nrOfIcons = 4;
                    break;
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    nrOfIcons = 5;
                    break;

                default:
                    throw new NotImplementedException("CreateBaseIconArr Can only handle Iconset types");
            };

            var arr = new ExcelConditionalFormattingIconDataBarValue[nrOfIcons];

            for (int i = 0; i < nrOfIcons; i++)
            {
                arr[i] = new ExcelConditionalFormattingIconDataBarValue
                    (eExcelConditionalFormattingValueObjectType.Percent, type);
            }

            return arr;
        }

        void ReadExtDxf(XmlReader xr)
        {

        }

        void ApplyIconSetExtValues(
            ExcelConditionalFormattingIconDataBarValue[] iconArr, 
            List<string> types, 
            List<double> values,
            List<string> customIconTypes = null,
            List<int> customIconIds = null)
        {
            for(int i = 0; i < iconArr.Length; i++)
            {
                iconArr[i].Type = types[i].ToEnum<eExcelConditionalFormattingValueObjectType>()
                    .GetValueOrDefault();

                iconArr[i].Value = values[i];

                if(customIconTypes != null)
                {
                    iconArr[i].SetCustomIconStringAndId(customIconTypes[i], customIconIds[i]);
                }
            }
        }

        Color GetColorFromExcelRgb(string rgb)
        {
            var colVal = int.Parse(rgb, NumberStyles.HexNumber);
            return Color.FromArgb(colVal);
        }

        string RemoveAuto(string typeString)
        {
            if(typeString.LastIndexOf("auto") == -1)
            {
                return typeString;
            }

            return typeString.Substring(typeString.LastIndexOf("auto"));
        }


        internal ExcelConditionalFormattingCollection(XmlReader xr, ExcelWorksheet ws)
        {
            _ws = ws;

            if(xr.LocalName == "conditionalFormattings")
            {
                ReadExtConditionalFormattings(xr);
            }
            else
            {
                while (xr.ReadUntil(1, "conditionalFormatting", "sheetData", "dataValidations", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks", "extLst", "pageMargins"))
                {
                    do
                    {
                        if (xr.LocalName == "conditionalFormatting" || xr.LocalName == "cfRule")
                        {
                            string address = null;

                            if (xr.LocalName == "conditionalFormatting")
                            {
                                address = xr.GetAttribute("sqref");
                            }
                            else
                            {
                                if(_rules.Count == 0)
                                {
                                    address = _dataBarStorage[_dataBarStorage.Count - 1].Address.Address;
                                }
                                else
                                {
                                    address = _rules[_rules.Count - 1].Address.Address;
                                }
                            }

                            if (address != null)
                            {
                                if (xr.NodeType == XmlNodeType.Element)
                                {
                                    if(xr.LocalName == "conditionalFormatting")
                                    {
                                        xr.Read();
                                    }

                                    var cf = ExcelConditionalFormattingRuleFactory.Create(new ExcelAddress(address), _ws, xr);

                                    if (cf.Type == eExcelConditionalFormattingRuleType.DataBar)
                                    {
                                        _dataBarStorage.Add(cf);
                                        _extLstDict.Add(((ExcelConditionalFormattingDataBar)cf).Uid, cf);
                                    }
                                    else
                                    {
                                        _rules.Add(cf);
                                    }
                                }
                                xr.Read();
                            }

                            //Handle many cfRules in one address
                            if (xr.LocalName != "cfRule")
                            {
                                xr.Read();
                            }
                        }
                    } while (xr.LocalName == "cfRule");
                }

                var adressLessCFs = new List<ExcelConditionalFormattingRule>();

                //identify ExtLst cfRules
                foreach (var cfRule in _rules)
                {
                    if (cfRule.IsExtLst)
                    {
                        if (cfRule.Type == eExcelConditionalFormattingRuleType.DataBar)
                        {
                            if(_extLstDict.ContainsKey(((ExcelConditionalFormattingDataBar)cfRule).Uid) == false)
                            {
                                _extLstDict.Add(((ExcelConditionalFormattingDataBar)cfRule).Uid, cfRule);
                            }
                        }
                        else
                        {
                            switch (cfRule.Type)
                            {
                                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                                    _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>)cfRule).Uid, cfRule);
                                    break;
                                case eExcelConditionalFormattingRuleType.FourIconSet:
                                    _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting4IconsSetType>)cfRule).Uid, cfRule);
                                    break;
                                case eExcelConditionalFormattingRuleType.FiveIconSet:
                                    _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>)cfRule).Uid, cfRule);
                                    break;
                                default:
                                    _extLstDict.Add(cfRule.Uid, cfRule);
                                    break;
                            }
                        }

                        if(string.IsNullOrEmpty(cfRule.Address.ToString()))
                        {
                            adressLessCFs.Add(cfRule);
                        }
                        else if(adressLessCFs.Count != 0)
                        {
                            foreach (var cf in adressLessCFs)
                            {
                                cf.Address = cfRule.Address;
                            }
                            adressLessCFs.Clear();
                        }
                    }
                }
            }
        }

        //Since a user could potentially change the type to and from an extType in iconSets?
        internal void UpdateExtDict()
        {
            _extLstDict.Clear();

            //identify ExtLst cfRules
            foreach (var cfRule in _rules)
            {
                if (cfRule.IsExtLst)
                {
                    switch (cfRule.Type)
                    {
                        case eExcelConditionalFormattingRuleType.ThreeIconSet:
                            _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>)cfRule).Uid, cfRule);
                            break;
                        case eExcelConditionalFormattingRuleType.FourIconSet:
                            _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting4IconsSetType>)cfRule).Uid, cfRule);
                            break;
                        case eExcelConditionalFormattingRuleType.FiveIconSet:
                            _extLstDict.Add(((ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>)cfRule).Uid, cfRule);
                            break;
                        case eExcelConditionalFormattingRuleType.DataBar:
                            _extLstDict.Add(((ExcelConditionalFormattingDataBar)cfRule).Uid, cfRule);
                            break;
                        case eExcelConditionalFormattingRuleType.TwoColorScale:
                            _extLstDict.Add(((ExcelConditionalFormattingTwoColorScale)cfRule).Uid, cfRule);
                            break;
                        default:
                            _extLstDict.Add(cfRule.Uid, cfRule);
                            break;
                    }
                }

                ////TODO: the sameAddressDict MUST be updated when users add addresses and must check 
                ////if anything outside of the dict has the address already
                //if(rulesOfSameAddressDict.ContainsKey(cfRule.Address))
                //{

                //}
            }

            foreach(var dataBar in _dataBarStorage)
            {
                _extLstDict.Add(((ExcelConditionalFormattingDataBar)dataBar).Uid, dataBar);
            }
        }

        internal void CopyRule(ExcelConditionalFormattingRule rule, ExcelAddress address = null)
        {
            var ruleCopy = rule.Clone();
            if (address != null)
            {
                ruleCopy.Address = address;
            }
            _rules.Add(ruleCopy);
        }

        IEnumerator<IExcelConditionalFormattingRule> IEnumerable<IExcelConditionalFormattingRule>.GetEnumerator()
        {
            for (int i = 0; i < _rules.Count; i++)
            {
                yield return _rules[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelConditionalFormattingRule this[int index]
        {
            get { return _rules[index]; }
            set { _rules[index] = value; }
        }

        /// <summary>
        /// Number of validations
        /// </summary>
        public int Count
        {
            get { return _rules.Count; }
        }

        /// <summary>
        /// Removes all 'cfRule' from the collection and from the XML.
        /// <remarks>
        /// This is the same as removing all the 'conditionalFormatting' nodes.
        /// </remarks>
        /// </summary>
        public void RemoveAll()
        {
            // Clear the <cfRule> item list
            _rules.Clear();
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its object
        /// </summary>
        /// <param name="item"></param>
        public void Remove(
          IExcelConditionalFormattingRule item)
        {
            Require.Argument(item).IsNotNull("item");

            try
            {
                _rules.Remove((ExcelConditionalFormattingRule)item);
            }
            catch
            {
                throw new Exception($"Cannot remove {item} as it is not part of this collection.");
            }
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its 0-based index
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(
          int index)
        {
            Require.Argument(index).IsInRange(0, this.Count - 1, "index");

            Remove(this[index]);
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its priority
        /// </summary>
        /// <param name="priority"></param>
        public void RemoveByPriority(
          int priority)
        {
            try
            {
                Remove(RulesByPriority(priority));
            }
            catch
            {
            }
        }

        /// <summary>
        /// Get a rule by its priority
        /// </summary>
        /// <param name="priority"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingRule RulesByPriority(
          int priority)
        {
            return _rules.Find(x => x.Priority == priority);
        }

        delegate ExcelConditionalFormattingRule Rule(ExcelAddress address, int priority, ExcelWorksheet ws);

        /// <summary>
        /// Add rule (internal)
        /// </summary>
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <param name="allowNullAddress"></param>
        /// <returns></returns>
        internal IExcelConditionalFormattingRule AddRule(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address, bool allowNullAddress = false)
        {
            if(!allowNullAddress)
            {
                Require.Argument(address).IsNotNull("address");
            }

            // Create the Rule according to the correct type, address and priority
            var cfRule = ExcelConditionalFormattingRuleFactory.Create(
              type,
              address,
              LastPriority++,
              _ws);

            // Add the newly created rule to the list
            _rules.Add(cfRule);

            UpdateExtDict();

            // Return the newly created rule
            return cfRule;
        }

        /// <summary>
        /// Add GreaterThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan(
          ExcelAddress address)
        {
            //var rule = new ExcelConditionalFormattingGreaterThan(address, LastPriority++, _ws);

            return (IExcelConditionalFormattingGreaterThan)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThan,
              address);
        }

        /// <summary>
        /// Add LessThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingLessThan AddLessThan(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingLessThan)AddRule(
              eExcelConditionalFormattingRuleType.LessThan,
              address);
        }

        /// <summary>
        /// Add between rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingBetween AddBetween(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingBetween)AddRule(
              eExcelConditionalFormattingRuleType.Between,
              address);
        }

        /// <summary>
        /// Add Equal rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingEqual AddEqual(ExcelAddress address)
        {
            return (IExcelConditionalFormattingEqual)AddRule(
              eExcelConditionalFormattingRuleType.Equal,
              address);
        }

        /// <summary>
        /// Add TextContains rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsText AddTextContains(ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              address);
        }

        /// <summary>
        /// Add Yesterday rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddYesterday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Yesterday,
              address);
        }

        /// <summary>
        /// Add Today rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddToday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Today,
              address);
        }

        /// <summary>
        /// Add Tomorrow rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Tomorrow,
              address);
        }

        /// <summary>
        /// Add Last7Days rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Last7Days,
              address);
        }

        /// <summary>
        /// Add lastWeek rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastWeek,
              address);
        }

        /// <summary>
        /// Add this week rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisWeek,
              address);
        }

        /// <summary>
        /// Add next week rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextWeek,
              address);
        }

        /// <summary>
        /// Add last month rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastMonth,
              address);
        }

        /// <summary>
        /// Add ThisMonth rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisMonth,
              address);
        }

        /// <summary>
        /// Add NextMonth rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextMonth,
              address);
        }

        /// <summary>
        /// Add DuplicateValues Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingDuplicateValues)AddRule(
              eExcelConditionalFormattingRuleType.DuplicateValues,
              address);
        }

        /// <summary>
        /// Add Bottom Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottom(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Bottom,
              address);
        }

        /// <summary>
        /// Add BottomPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.BottomPercent,
              address);
        }

        /// <summary>
        /// Add Top Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTop(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Top,
              address);
        }

        /// <summary>
        /// Add TopPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTopPercent(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.TopPercent,
              address);
        }

        /// <summary>
        /// Add AboveAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveAverage,
              address);  
        }

        /// <summary>
        /// Add AboveOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
              address);
        }

        /// <summary>
        /// Add BelowAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowAverage,
              address);
        }

        /// <summary>
        /// Add BelowOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
              address);
        }

        /// <summary>
        /// Add AboveStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveStdDev,
              address);
        }

        /// <summary>
        /// Add BelowStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowStdDev,
              address);
        }

        /// <summary>
        /// Add BeginsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingBeginsWith AddBeginsWith(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingBeginsWith)AddRule(
              eExcelConditionalFormattingRuleType.BeginsWith,
              address);
        }

        /// <summary>
        /// Add ContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.ContainsBlanks,
              address);
        }

        /// <summary>
        /// Add ContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsErrors AddContainsErrors(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.ContainsErrors,
              address);
        }

        /// <summary>
        /// Add ContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsText AddContainsText(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              address);
        }

        /// <summary>
        /// Add EndsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingEndsWith AddEndsWith(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingEndsWith)AddRule(
              eExcelConditionalFormattingRuleType.EndsWith,
              address);
        }

        /// <summary>
        /// Add Expression Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingExpression AddExpression(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingExpression)AddRule(
              eExcelConditionalFormattingRuleType.Expression,
              address);
        }

        /// <summary>
        /// Add GreaterThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
              address);
        }

        /// <summary>
        /// Add LessThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingLessThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.LessThanOrEqual,
              address);
        }

        /// <summary>
        /// Add NotBetween Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotBetween AddNotBetween(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotBetween)AddRule(
              eExcelConditionalFormattingRuleType.NotBetween,
              address);
        }

        /// <summary>
        /// Add NotContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsBlanks,
              address);
        }

        /// <summary>
        /// Add NotContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsErrors,
              address);
        }

        /// <summary>
        /// Add NotContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsText AddNotContainsText(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsText)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsText,
              address);
        }

        /// <summary>
        /// Add NotEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotEqual AddNotEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotEqual)AddRule(
              eExcelConditionalFormattingRuleType.NotEqual,
              address);
        }

        /// <summary>
        /// Add Unique Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingUniqueValues AddUniqueValues(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingUniqueValues)AddRule(
              eExcelConditionalFormattingRuleType.UniqueValues,
              address);
        }

        /// <summary>
        /// Add ThreeColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingThreeColorScale)AddRule(
              eExcelConditionalFormattingRuleType.ThreeColorScale,
              address);
        }

        /// <summary>
        /// Add TwoColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTwoColorScale)AddRule(
              eExcelConditionalFormattingRuleType.TwoColorScale,
              address);
        }

        /// <summary>
        /// Add ThreeIconSet Rule
        /// </summary>
        /// <param name="Address">The address</param>
        /// <param name="IconSet">Type of iconset</param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(ExcelAddress Address, eExcelconditionalFormatting3IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.ThreeIconSet,
                Address);
            icon.IconSet = IconSet;

            UpdateExtDict();
            return icon;
        }
        /// <summary>
        /// Adds a FourIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(ExcelAddress Address, eExcelconditionalFormatting4IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.FourIconSet,
                Address);
            icon.IconSet = IconSet;
            UpdateExtDict();

            return icon;
        }
        /// <summary>
        /// Adds a FiveIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(ExcelAddress Address, eExcelconditionalFormatting5IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFiveIconSet)AddRule(
                eExcelConditionalFormattingRuleType.FiveIconSet,
                Address);
            icon.IconSet = IconSet;
            UpdateExtDict();

            return icon;
        }
        /// <summary>
        /// Adds a databar rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDataBarGroup AddDatabar(ExcelAddress Address, Color color)
        {
            var dataBar = (IExcelConditionalFormattingDataBarGroup)AddRule(
                eExcelConditionalFormattingRuleType.DataBar,
                Address);
            dataBar.Color = color;
            UpdateExtDict();

            return dataBar;
        }
    }
}
