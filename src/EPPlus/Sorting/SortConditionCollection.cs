/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sorting
{
    public class SortConditionCollection : XmlHelper, IEnumerable<SortCondition>
    {
        internal SortConditionCollection(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            var conditionNodes = topNode.SelectNodes("//d:sortCondition", nameSpaceManager);
            if(conditionNodes != null)
            {
                foreach(var node in conditionNodes)
                {
                    var condition = new SortCondition(nameSpaceManager, (XmlNode)node);
                    _sortConditions.Add(condition);
                }
            }
        }

        private readonly List<SortCondition> _sortConditions = new List<SortCondition>();
        private readonly string _sortConditionPath = "d:sortCondition";

        public IEnumerator<SortCondition> GetEnumerator()
        {
            return _sortConditions.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _sortConditions.GetEnumerator();
        }

        internal void Add(string @ref, bool? decending = null)
        {
            if (_sortConditions.Count > 63) throw new ArgumentException("Too many sort conditions added, max number of conditions is 64");
            var node = CreateNode(TopNode, _sortConditionPath, true);
            var condition = new SortCondition(NameSpaceManager, node);
            condition.Ref = @ref;
            if(decending.HasValue)
            {
                condition.Descending = decending.Value;
            }
            TopNode.AppendChild(condition.TopNode);
            _sortConditions.Add(condition);
        }

        internal void Add(string @ref, bool? decending, string[] customList = null)
        {
            if (_sortConditions.Count > 63) throw new ArgumentException("Too many sort conditions added, max number of conditions is 64");
            var node = CreateNode(TopNode, _sortConditionPath, true);
            var condition = new SortCondition(NameSpaceManager, node);
            condition.Ref = @ref;
            if (decending.HasValue)
            {
                condition.Descending = decending.Value;
            }
            condition.CustomList = customList;
            TopNode.AppendChild(condition.TopNode);
            _sortConditions.Add(condition);
        }
    }
}
