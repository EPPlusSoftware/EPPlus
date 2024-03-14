/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
{
    internal class CssRuleCollection : IEnumerable<CssRule>
    {
        List<CssRule> _cssRules;

        internal List<CssRule> CssRules => _cssRules;

        internal CssRuleCollection()
        {
            _cssRules = new List<CssRule>();
        }

        IEnumerator<CssRule> IEnumerable<CssRule>.GetEnumerator()
        {
            for (int i = 0; i < _cssRules.Count; i++)
            {
                yield return _cssRules[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _cssRules.GetEnumerator();
        }

        internal void AddRule(string ruleName, string declarationName, params string[] declarationValues)
        {
            var toBeAdded = new CssRule(ruleName)
            {
                Declarations =
                {
                    new Declaration(declarationName, declarationValues),
                }
            };

            _cssRules.Add(toBeAdded);
        }

        internal void AddRule(CssRule rule)
        {
            _cssRules.Add(rule);
        }

        internal void RemoveRule(CssRule rule)
        {
            _cssRules.Remove(rule);
        }

        internal void RemoveRuleByName(string ruleName)
        {
            _cssRules.RemoveAll(x => x.Selector == ruleName);
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public CssRule this[int index]
        {
            get { return _cssRules[index]; }
            set { _cssRules[index] = value; }
        }
    }
}
