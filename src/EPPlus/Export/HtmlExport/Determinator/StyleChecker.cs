using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Determinator
{
    internal class StyleChecker
    {
        ExcelStyles _wbStyles;
        IStyleExport _style = null;
        int _id = -1;

        internal int Id => _id;

        List<IStyleExport> _styleList;

        internal StyleChecker(ExcelStyles wbStyles)
        {
            _wbStyles = wbStyles;
        }

        internal StyleCache Cache { private get; set; } = null;

        internal IStyleExport Style 
        {
            private get
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
            if (Style == null || Cache == null)
            {
                throw new InvalidOperationException("Must assign Style and Cache to Stylechecker first");
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

            //var delegator = new ForEachCFDelegator(dict);

            //delegator.FuncOnEachElement(cellAddress, testStyle);
            //if (ShouldAdd)
            //{
            //    addToCss(Id, _styleList);
            //}
        }

        bool testStyle(ExcelConditionalFormattingRule cf)
        {
            _style = new StyleDxf(cf.Style);
            return true;
        }

        //if (ce.CellAddress != null && _cfAtAddresses.ContainsKey(ce.CellAddress))
        //{
        //    foreach (var cf in _cfAtAddresses[ce.CellAddress])
        //    {
        //        var dxfStyle = new StyleDxf(cf._style);
        //        ScDxf.Style = dxfStyle;

        //        if (!ScDxf.ShouldAdd)
        //        {
        //            cssTranslator.AddToCollection(sc.GetStyleList(), styles.GetNormalStyle(), ScDxf.Id);
        //        }
        //    }
        //}

    }
}
