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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Determinator
{
    internal class StyleChecker
    {
        ExcelStyles _wbStyles;
        IStyleExport _style = null;
        int _id = -1;

        internal int Id => _id;

        List<IStyleExport> _styleList;

        internal StyleChecker(ExcelStyles wbStyles, IStyleExport style , StyleCache cache)
        {
            _wbStyles = wbStyles;
            _styleList = new List<IStyleExport>();
            Style = style;
            Cache = cache;
        }

        private StyleCache Cache { get; set; } = null;

        private IStyleExport Style 
        {
            get
            {
                return _style;
            }
            set
            {
                _style = value;
                _styleList.Clear();

                _styleList.Add(value);
            }
        }

        internal bool IsAdded(int bottomStyleId = -1, int rightStyleId = -1)
        {
            if (Style.HasStyle == false)
            {
                return false;
            }

            var key = Style.StyleKey;
            if (bottomStyleId > -1) key += bottomStyleId + "|" + rightStyleId;

            return Cache.IsAdded(key, out _id);
        }

        internal bool BorderStyleCheck(int borderIdBottom, int borderIdRight)
        {
            return (Style.HasStyle || borderIdBottom > 0 || borderIdRight > 0);
        }

        internal bool ShouldAdd 
        { 
            get
            {
                //If already added we should 'not' add
                return !IsAdded();
            }
        }

        internal bool ShouldAddWithBorders(int bottomStyleId, int rightStyleId)
        {
            if (IsAdded(bottomStyleId, rightStyleId))
            {
                return false;
            }

            _styleList.Add(new StyleXml(_wbStyles.CellXfs[bottomStyleId]));
            _styleList.Add(new StyleXml(_wbStyles.CellXfs[rightStyleId]));

            return BorderStyleCheck(_wbStyles.CellXfs[bottomStyleId].BorderId, _wbStyles.CellXfs[rightStyleId].BorderId);
        }

        internal List<IStyleExport> GetStyleList()
        {
            return _styleList;
        }

        internal void AddConditionalFormattingsToCollection(string cellAddress, Dictionary<string, List<ExcelConditionalFormattingRule>> dict, Func<int, List<IStyleExport>, bool> addToCss)
        {
            if (cellAddress != null && dict.ContainsKey(cellAddress))
            {
                foreach (var cf in dict[cellAddress])
                {
                    _style = new StyleDxf(cf.Style);
                    if (ShouldAdd)
                    {
                        addToCss(Id, _styleList);
                    }
                }
            }
        }
    }
}
