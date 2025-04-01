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
using OfficeOpenXml.Core.RangeQuadTree;
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
    /// <summary>
    /// Collection of all ConditionalFormattings in the workbook
    /// </summary>
    public class ExcelConditionalFormattingCollection : IEnumerable<IExcelConditionalFormattingRule>
    {
        List<ExcelConditionalFormattingRule> _rules = new List<ExcelConditionalFormattingRule>();
        ExcelWorksheet _ws;
        int LastPriority = 1;
        //A dict for those conditionalFormattings that are Ext, have been read in locally but not yet read in their ExtLst parts.
        internal Dictionary<string, ExcelConditionalFormattingRule> localAndExtDict = new Dictionary<string, ExcelConditionalFormattingRule>();

        internal ExcelConditionalFormattingCollection(ExcelWorksheet ws)
        {
            _ws = ws;
            _rules = new List<ExcelConditionalFormattingRule>();
            if(_ws.Dimension==null)
            {
                CfIndex = new QuadTree<IExcelConditionalFormattingRule>();
            }
            else
            {
                CfIndex = new QuadTree<IExcelConditionalFormattingRule>(_ws.Dimension);
            }
        }
        internal QuadTree<IExcelConditionalFormattingRule> CfIndex { get; set; }

        internal void ReadRegularConditionalFormattings(XmlReader xr)
        {
            string address = null;
			bool pivot = false;
            while (xr.ReadUntil(1, "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst" ))
            {
                address = null;

                do
                {
                    if (xr.LocalName == "conditionalFormatting")
                    {
                        address = xr.GetAttribute("sqref");
                        if (string.IsNullOrEmpty(address) == false)
                        {
							address = address.Replace(' ', ',');
						}

						pivot = xr.GetAttribute("pivot") == "1";

                        xr.Read();
                    }

                    if (xr.LocalName == "cfRule" && xr.NodeType == XmlNodeType.Element)
                    {
                        ExcelConditionalFormattingRule cf;

						if (string.IsNullOrEmpty(address))
                        {
							cf = ExcelConditionalFormattingRuleFactory.Create(null, _ws, xr);
						}
						else
                        {
                            cf = ExcelConditionalFormattingRuleFactory.Create(new ExcelAddress(address), _ws, xr);
                        }
						cf.PivotTable = pivot;
						//If cf exists in both local and ExtLst spaces
						if (cf.IsExtLst && cf._uid != null)
						{
							localAndExtDict.Add(cf._uid.Trim('{','}'), cf);
						}
						else
						{
							AddToList(cf);
						}
					}
                    while ((xr.LocalName == "conditionalFormatting" || xr.LocalName == "cfRule") && xr.NodeType == XmlNodeType.EndElement) xr.Read();
                }
                while (xr.LocalName == "conditionalFormatting" || xr.LocalName == "cfRule");
            }
        }

        /// <summary>
        /// Read conditionalFormatting info from extLst in xml via xr reader
        /// </summary>
        internal void ReadExtConditionalFormattings(XmlReader xr)
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
                        string id = xr.GetAttribute("id").Trim('{','}');

                        if (string.IsNullOrEmpty(id))
                        {
                            throw new InvalidOperationException("XML invalid. cfRule without Id found");
                        }

                        if (xr.GetAttribute("type") == "dataBar")
                        {
                            //cfRule->Type
                            xr.Read();

                            var dataBar = (ExcelConditionalFormattingDataBar)localAndExtDict[id];
                            dataBar.LowValue.minLength = int.Parse(xr.GetAttribute("minLength"));
                            dataBar.HighValue.maxLength = int.Parse(xr.GetAttribute("maxLength"));

                            dataBar.ShowValue = string.IsNullOrEmpty(xr.GetAttribute("showValue")) ? true : xr.GetAttribute("showValue") != "0";
                            dataBar.Border = string.IsNullOrEmpty(xr.GetAttribute("border")) ? false : xr.GetAttribute("border") != "0";
                            dataBar.Gradient = string.IsNullOrEmpty(xr.GetAttribute("gradient")) ? true : xr.GetAttribute("gradient") != "0";

                            bool? negativeBarBorderColorSameAsPositive = null;
                            bool? negativeBarColorSameAsPositive = null;


                            if (!string.IsNullOrEmpty(xr.GetAttribute("direction")))
                            {
                                dataBar.Direction = (eDatabarDirection)xr.GetAttribute("direction").ToEnum<eDatabarDirection>();   
                            }

                            if(!string.IsNullOrEmpty(xr.GetAttribute("negativeBarBorderColorSameAsPositive")))
                            {
                                negativeBarBorderColorSameAsPositive = xr.GetAttribute("negativeBarBorderColorSameAsPositive") != "0";
                            }

                            if (!string.IsNullOrEmpty(xr.GetAttribute("negativeBarColorSameAsPositive")))
                            {
                                negativeBarColorSameAsPositive = xr.GetAttribute("negativeBarColorSameAsPositive") != "0";
                            }

                            if (!string.IsNullOrEmpty(xr.GetAttribute("axisPosition")))
                            {
                                dataBar.AxisPosition = (eExcelDatabarAxisPosition)xr.GetAttribute("axisPosition").ToEnum<eExcelDatabarAxisPosition>();
                            }

                            //CfRule -> cfvo
                            xr.Read();

                            string typeString1 = RemoveAuto(xr.GetAttribute("type"));

                            dataBar.LowValue.Type = typeString1.ToEnum<eExcelConditionalFormattingValueObjectType>().Value;

                            xr.Read();

                            if (dataBar.LowValue.HasValueOrFormula && xr.Name == "xm:f")
                            {
                                xr.Read();
                                var content = xr.ReadContentAsString();
                                if (double.TryParse(content, NumberStyles.Any, CultureInfo.InvariantCulture, out double result)
                                    && dataBar.LowValue.Type != eExcelConditionalFormattingValueObjectType.Formula)
                                {
                                    dataBar.LowValue.Value = result;
                                }
                                else
                                {
                                    dataBar.LowValue.Formula = content;
                                }
                                xr.Read();
                                xr.Read();
                            }

                            string typeString2 = RemoveAuto(xr.GetAttribute("type"));

                            dataBar.HighValue.Type = typeString2.ToEnum<eExcelConditionalFormattingValueObjectType>().Value;

                            xr.Read();

                            if (dataBar.HighValue.HasValueOrFormula && xr.Name == "xm:f")
                            {
                                xr.Read();
                                var content = xr.ReadContentAsString();
                                if (double.TryParse(content, NumberStyles.Any, CultureInfo.InvariantCulture, out double result)
                                    && dataBar.HighValue.Type != eExcelConditionalFormattingValueObjectType.Formula)
                                {
                                    dataBar.HighValue.Value = result;
                                }
                                else
                                {
                                    dataBar.HighValue.Formula = content;
                                }
                                xr.Read();
                                xr.Read();
                            }

                            dataBar.ReadInCTColor(xr);

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

                            if(negativeBarBorderColorSameAsPositive != null)
                            {
                                dataBar.NegativeBarBorderColorSameAsPositive = negativeBarBorderColorSameAsPositive.Value;
                            }

                            if(negativeBarColorSameAsPositive != null)
                            {
                                dataBar.NegativeBarColorSameAsPositive = negativeBarColorSameAsPositive.Value;
                            }

                            AddToList(dataBar);
                            dataBar.Uid = id;
                        }
                        else if (xr.GetAttribute("type") == "iconSet")
                        {
                            int priority = int.Parse(xr.GetAttribute("priority"));

                            //cfRule->Type
                            xr.Read();

                            string iconSet = xr.GetAttribute("iconSet");

                            bool showValue = string.IsNullOrEmpty(xr.GetAttribute("showValue")) ? true : xr.GetAttribute("showValue") != "0";
                            bool percent = string.IsNullOrEmpty(xr.GetAttribute("percent")) ? true : xr.GetAttribute("percent") != "0";
                            bool reverse = string.IsNullOrEmpty(xr.GetAttribute("reverse")) ? false : xr.GetAttribute("reverse") != "0";
                            bool custom = string.IsNullOrEmpty(xr.GetAttribute("custom")) ? false : xr.GetAttribute("custom") != "0";

                            xr.Read();

                            var types = new List<string>();
                            var values = new List<string>();
                            var gteValues = new List<bool>();

                            do
                            {
                                types.Add(xr.GetAttribute("type"));
                                var test = xr.GetAttribute("gte");
                                gteValues.Add(test != "0");

                                xr.Read();
                                xr.Read();

                                values.Add(xr.Value);

                                xr.Read();
                                xr.Read();
                                xr.Read();
                            } while (xr.LocalName == "cfvo");

                            var numIcons = types.Count();

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

                            //iconSet->cfRule->sqref
                            string address = null;
                            xr.Read();
                            xr.Read();

                            if (xr.LocalName != "cfRule")
                            {
                                xr.Read();

                                address = xr.ReadContentAsString();

                                xr.Read();
                            }

                            ExcelAddress iconAddress = null;
                            if (address != null)
                            {
                                iconAddress = new ExcelAddress(address);
                            }

                            ExcelConditionalFormattingRule rule = null;

                            switch (numIcons.ToString()[0])
                            {
                                case '3':

                                    IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> threeIconSet;

                                    threeIconSet = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)
                                        AddRule(eExcelConditionalFormattingRuleType.ThreeIconSet, iconAddress, true);

                                    if (iconSet != null)
                                    {
                                        threeIconSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting3IconsSetType>().Value;
                                    }

                                    ApplyIconSetExtValues(
                                        new ExcelConditionalFormattingIconDataBarValue[]
                                        { threeIconSet.Icon1, threeIconSet.Icon2, threeIconSet.Icon3 },
                                        types, values, gteValues, customIconTypes, customIconIds);

                                    ApplyIconSetAttributes(showValue, percent, reverse, threeIconSet);

                                    rule = (ExcelConditionalFormattingRule)threeIconSet;

                                    break;

                                case '4':

                                    IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> fourSet;

                                    fourSet = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)
                                        AddRule(eExcelConditionalFormattingRuleType.FourIconSet, iconAddress, true);

                                    if (iconSet != null)
                                    {
                                        fourSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting4IconsSetType>().Value;
                                    }

                                    ApplyIconSetExtValues(
                                    new ExcelConditionalFormattingIconDataBarValue[]
                                    { fourSet.Icon1, fourSet.Icon2, fourSet.Icon3, fourSet.Icon4 },
                                    types, values, gteValues, customIconTypes, customIconIds);

                                    ApplyIconSetAttributes(showValue, percent, reverse, fourSet);

                                    rule = (ExcelConditionalFormattingRule)fourSet;

                                    break;

                                case '5':
                                    var fiveSet = (IExcelConditionalFormattingFiveIconSet)
                                        AddRule(eExcelConditionalFormattingRuleType.FiveIconSet, iconAddress, true);

                                    if (iconSet != null)
                                    {
                                        fiveSet.IconSet = iconSet.Substring(1).ToEnum<eExcelconditionalFormatting5IconsSetType>().Value;
                                    }

                                    ApplyIconSetExtValues(
                                     new ExcelConditionalFormattingIconDataBarValue[]
                                     { fiveSet.Icon1, fiveSet.Icon2, fiveSet.Icon3, fiveSet.Icon4 , fiveSet.Icon5 },
                                     types, values, gteValues, customIconTypes, customIconIds);

                                    ApplyIconSetAttributes(showValue, percent, reverse, fiveSet);

                                    rule = (ExcelConditionalFormattingRule)fiveSet;
                                    break;
                            }

                            rule.Priority = priority;
                            rule.Uid = id;

                            if (iconAddress == null && rule != null)
                            {
                                addresslessCFs.Add(rule);
                            }
                        }
                        else
                        {
                            var cf = ExcelConditionalFormattingRuleFactory.Create(null, _ws, xr);
                            cf.Uid = id;
                            AddToList(cf);

                            if (cf.Address == null)
                            {
                                addresslessCFs.Add(cf);
                            }
                        }
                    } while (xr.LocalName == "cfRule");

                    var latestAddress = _rules.LastOrDefault().Address;

                    if (xr.LocalName == "sqref" && xr.NodeType != XmlNodeType.EndElement)
                    {
                        xr.Read();
                        latestAddress = new ExcelAddress(xr.ReadString());
                    }

                    foreach (var cf in addresslessCFs)
                    {
                        cf.Address = latestAddress;
                    }
                }
            }
        }

        void ApplyIconSetAttributes<T>(bool showValue, bool percent, bool reverse, IExcelConditionalFormattingIconSetGroup<T> group)
        {
            group.ShowValue = showValue;
            group.IconSetPercent = percent;
            group.Reverse = reverse;
        }

        void ApplyIconSetExtValues(
            ExcelConditionalFormattingIconDataBarValue[] iconArr, 
            List<string> types, 
            List<string> values,
            List<bool> gteValues,
            List<string> customIconTypes = null,
            List<int> customIconIds = null)
        {
            for(int i = 0; i < iconArr.Length; i++)
            {
                iconArr[i].Type = types[i].ToEnum<eExcelConditionalFormattingValueObjectType>()
                    .GetValueOrDefault();

                if(double.TryParse(values[i], out double result))
                {
                    iconArr[i].Value = result;
                }
                else
                {
                    iconArr[i].Formula = values[i];
                }

                if (gteValues[i] == false)
                {
                    iconArr[i].GreaterThanOrEqualTo = gteValues[i];
                }

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

        internal void CopyRule(ExcelConditionalFormattingRule rule, ExcelAddress address = null)
        {
            ExcelConditionalFormattingRule ruleCopy = null;

            if (rule._ws != _ws)
            {
                ruleCopy = rule.Clone(_ws);
            }
            else
            {
                ruleCopy = rule.Clone();
            }

            if (address != null)
            {
                ruleCopy.Address = address;
            }
            AddToList(ruleCopy);
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
                CfIndex.Clear(item.Address, item);
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
            catch(Exception ex)
            {
                throw new InvalidOperationException($"Could not remove item with priority {priority}", ex);
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

        internal void ChangePriority(ExcelConditionalFormattingRule rule, int priorityNew)
        {
            if (RulesByPriority(priorityNew) != null)
            {
                if (rule.Priority > priorityNew)
                {
                    for (int i = rule.Priority - 1; i >= priorityNew; i--)
                    {
                        var cfRule = (ExcelConditionalFormattingRule)RulesByPriority(i);

                        if (cfRule != null)
                        {
                            cfRule._priority++;
                        }
                    }
                }
                else
                {
                    for (int i = rule.Priority + 1; i <= priorityNew; i++)
                    {
                        var cfRule = (ExcelConditionalFormattingRule)RulesByPriority(i);

                        if (cfRule != null)
                        {
                            cfRule._priority--;
                        }
                    }
                }
            }
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
            if (!allowNullAddress)
            {
                Require.Argument(address).IsNotNull("address");

                if (!ExcelCellBase.IsValidAddress(address.Address))
                {
                    throw new ArgumentException(
                        $"Address: \"{address.Address}\" for Conditional Formatting rule of type {type} " +
                        $"is not a valid address");
                }
            }

            // Create the Rule according to the correct type, address and priority
            var cfRule = ExcelConditionalFormattingRuleFactory.Create(
              type,
              address,
              LastPriority++,
              _ws);

            // Add the newly created rule to the list
            AddToList(cfRule);

            // Return the newly created rule
            return cfRule;
        }

        private void AddToList(ExcelConditionalFormattingRule cfRule)
        {
            _rules.Add(cfRule);
        }

        internal void ClearTempExportCacheForAllCFs()
        {
            foreach(var cf in _rules)
            {
                cf.RemoveTempExportData();
            }
        }

        /// <summary>
        /// Add GreaterThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingGreaterThan)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThan,
              address);
        }

        /// <summary>
        /// Add GreaterThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan(
          string address)
        {
            return (IExcelConditionalFormattingGreaterThan)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThan,
              new ExcelAddress(address));
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
        /// Add LessThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingLessThan AddLessThan(
            string address)
        {
            return (IExcelConditionalFormattingLessThan)AddRule(
              eExcelConditionalFormattingRuleType.LessThan,
              new ExcelAddress(address));
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
        /// Add between rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingBetween AddBetween(
            string address)
        {
            return (IExcelConditionalFormattingBetween)AddRule(
              eExcelConditionalFormattingRuleType.Between,
              new ExcelAddress(address));
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
        /// Add Equal rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingEqual AddEqual(string address)
        {
            return (IExcelConditionalFormattingEqual)AddRule(
              eExcelConditionalFormattingRuleType.Equal,
            new ExcelAddress(address));
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
        /// Add TextContains rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsText AddTextContains(string address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              new ExcelAddress(address));
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
        /// Add Yesterday rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddYesterday(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Yesterday,
              new ExcelAddress(address));
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
        /// Add Today rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddToday(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Today,
              new ExcelAddress(address));
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
        /// Add Tomorrow rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Tomorrow,
              new ExcelAddress(address));
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
        /// Add Last7Days rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Last7Days,
              new ExcelAddress(address));
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
        /// Add lastWeek rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastWeek,
              new ExcelAddress(address));
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
        /// Add this week rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisWeek,
              new ExcelAddress(address));
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
        /// Add next week rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextWeek,
              new ExcelAddress(address));
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
        /// Add last month rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastMonth,
              new ExcelAddress(address));
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
        /// Add ThisMonth rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisMonth,
              new ExcelAddress(address));
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
        /// Add NextMonth rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(string address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextMonth,
              new ExcelAddress(address));
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
        /// Add DuplicateValues Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(
         string address)
        {
            return (IExcelConditionalFormattingDuplicateValues)AddRule(
              eExcelConditionalFormattingRuleType.DuplicateValues,
              new ExcelAddress(address));
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
        /// Add Bottom Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottom(
          string address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Bottom,
              new ExcelAddress(address));
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
        /// Add BottomPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(
          string address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.BottomPercent,
              new ExcelAddress(address));
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
        /// Add Top Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTop(
          string address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Top,
              new ExcelAddress(address));
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
        /// Add TopPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTopPercent(
          string address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.TopPercent,
              new ExcelAddress(address));
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
        /// Add AboveAverage Rule
        /// </summary>
        /// <param name="address">String must be a valid excelAddress</param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveAverage(
          string address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveAverage,
              new ExcelAddress(address));
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
        /// Add AboveOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(
          string address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
              new ExcelAddress(address));
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
        /// Add BelowAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowAverage(
          string address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowAverage,
              new ExcelAddress(address));
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
        /// Add BelowOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(
          string address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
              new ExcelAddress(address));
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
        /// Add AboveStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(
          string address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveStdDev,
              new ExcelAddress(address));
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
        /// Add BelowStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(
          string address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowStdDev,
              new ExcelAddress(address));
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
        /// Add BeginsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingBeginsWith AddBeginsWith(
          string address)
        {
            return (IExcelConditionalFormattingBeginsWith)AddRule(
              eExcelConditionalFormattingRuleType.BeginsWith,
              new ExcelAddress(address));
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
        /// Add ContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(
          string address)
        {
            return (IExcelConditionalFormattingContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.ContainsBlanks,
              new ExcelAddress(address));
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
        /// Add ContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsErrors AddContainsErrors(
          string address)
        {
            return (IExcelConditionalFormattingContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.ContainsErrors,
              new ExcelAddress(address));
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
        /// Add ContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsText AddContainsText(
          string address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              new ExcelAddress(address));
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
        /// Add EndsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingEndsWith AddEndsWith(
          string address)
        {
            return (IExcelConditionalFormattingEndsWith)AddRule(
              eExcelConditionalFormattingRuleType.EndsWith,
              new ExcelAddress(address));
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
        /// Add Expression Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingExpression AddExpression(
          string address)
        {
            return (IExcelConditionalFormattingExpression)AddRule(
              eExcelConditionalFormattingRuleType.Expression,
              new ExcelAddress(address));
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
        /// Add GreaterThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(
          string address)
        {
            return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
              new ExcelAddress(address));
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
        /// Add LessThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(
          string address)
        {
            return (IExcelConditionalFormattingLessThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.LessThanOrEqual,
              new ExcelAddress(address));
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
        /// Add NotBetween Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotBetween AddNotBetween(
          string address)
        {
            return (IExcelConditionalFormattingNotBetween)AddRule(
              eExcelConditionalFormattingRuleType.NotBetween,
              new ExcelAddress(address));
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
        /// Add NotContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(
          string address)
        {
            return (IExcelConditionalFormattingNotContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsBlanks,
              new ExcelAddress(address));
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
        /// Add NotContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(
          string address)
        {
            return (IExcelConditionalFormattingNotContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsErrors,
              new ExcelAddress(address));
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
        /// Add NotContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsText AddNotContainsText(
          string address)
        {
            return (IExcelConditionalFormattingNotContainsText)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsText,
              new ExcelAddress(address));
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
        /// Add NotEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotEqual AddNotEqual(
          string address)
        {
            return (IExcelConditionalFormattingNotEqual)AddRule(
              eExcelConditionalFormattingRuleType.NotEqual,
              new ExcelAddress(address));
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
        /// Add Unique Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingUniqueValues AddUniqueValues(
          string address)
        {
            return (IExcelConditionalFormattingUniqueValues)AddRule(
              eExcelConditionalFormattingRuleType.UniqueValues,
              new ExcelAddress(address));
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
        /// Add ThreeColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(
          string address)
        {
            return (IExcelConditionalFormattingThreeColorScale)AddRule(
              eExcelConditionalFormattingRuleType.ThreeColorScale,
              new ExcelAddress(address));
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
        /// Add TwoColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(
          string address)
        {
            return (IExcelConditionalFormattingTwoColorScale)AddRule(
              eExcelConditionalFormattingRuleType.TwoColorScale,
              new ExcelAddress(address));
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

            return icon;
        }

        /// <summary>
        /// Add ThreeIconSet Rule
        /// </summary>
        /// <param name="Address">The address</param>
        /// <param name="IconSet">Type of iconset</param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(string Address, eExcelconditionalFormatting3IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.ThreeIconSet,
                new ExcelAddress(Address));
            icon.IconSet = IconSet;

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

            return icon;
        }

        /// <summary>
        /// Adds a FourIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(string Address, eExcelconditionalFormatting4IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.FourIconSet,
                new ExcelAddress(Address));
            icon.IconSet = IconSet;

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

            return icon;
        }

        /// <summary>
        /// Adds a FiveIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(string Address, eExcelconditionalFormatting5IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFiveIconSet)AddRule(
                eExcelConditionalFormattingRuleType.FiveIconSet,
                new ExcelAddress(Address));
            icon.IconSet = IconSet;

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

            return dataBar;
        }

        /// <summary>
        /// Adds a databar rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDataBarGroup AddDatabar(string Address, Color color)
        {
            var dataBar = (IExcelConditionalFormattingDataBarGroup)AddRule(
                eExcelConditionalFormattingRuleType.DataBar,
                new ExcelAddress(Address));
            dataBar.Color = color;
            dataBar.BorderColor.Color = color;

            return dataBar;
        }

        internal IExcelConditionalFormattingRule GetByPriority(int priority)
        {
            foreach (var rule in _rules)
            {
                if(rule.Priority == priority)
                {
                    return rule;
                }
            }
            return null;
        }

        internal List<QuadRangeItem<IExcelConditionalFormattingRule>> GetIntersectingRanges(ExcelAddress address)
        {
            return CfIndex.GetIntersectingRangeItems(new QuadRange(address));
        }
    }
}
