using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Writers.Css
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
