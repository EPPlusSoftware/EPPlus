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
 *************************************************************************************************/
using System;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using draw=System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Slicer.Style;
using OfficeOpenXml.Table;
using System.Globalization;

namespace OfficeOpenXml
{
    /// <summary>
    /// Containts all shared cell styles for a workbook
    /// </summary>
    public sealed class ExcelStyles : XmlHelper
    {
        const string ColorsPath = "d:colors/d:indexedColors/d:rgbColor";
        const string NumberFormatsPath = "d:numFmts";
        const string FontsPath = "d:fonts";
        const string FillsPath = "d:fills";
        const string BordersPath = "d:borders";
        const string CellStyleXfsPath = "d:cellStyleXfs";
        const string CellXfsPath = "d:cellXfs";
        const string CellStylesPath = "d:cellStyles";
        const string TableStylesPath = "d:tableStyles";
        internal const string DxfsPath = "d:dxfs";
        internal const string DxfSlicerStylesPath = "d:extLst/d:ext[@uri='" + ExtLstUris.SlicerStylesDxfCollectionUri + "']/x14:dxfs";
        const string SlicerStylesPath = "d:extLst/d:ext[@uri='" + ExtLstUris.SlicerStylesUri + "']/x14:slicerStyles";
        XmlDocument _styleXml;
        internal ExcelWorkbook _wb;
        ExcelNamedStyleXml _normalStyle;
        XmlNamespaceManager _nameSpaceManager;
        internal int _nextDfxNumFmtID = 164;

        internal ExcelStyles(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml.DocumentElement)
        {
            _styleXml = xml;
            _wb = wb;
            _nameSpaceManager = NameSpaceManager;
            SchemaNodeOrder = new string[] { "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs" };
            LoadFromDocument();
        }
        /// <summary>
        /// Loads the style XML to memory
        /// </summary>
        private void LoadFromDocument()
        {
            //colors
            XmlNodeList colorNodes = GetNodes(ColorsPath);
            if (colorNodes != null && colorNodes.Count > 0)
            {
                int index = 0;
                foreach (XmlNode node in colorNodes)
                {
                    ExcelColor.indexedColors[index++] = "#" + node.Attributes["rgb"].InnerText;
                }
            }

            //NumberFormats
            ExcelNumberFormatXml.AddBuildIn(NameSpaceManager, NumberFormats);
            XmlNode numNode = GetNode(NumberFormatsPath);
            if (numNode != null)
            {
                foreach (XmlNode n in numNode)
                {
                    ExcelNumberFormatXml nf = new ExcelNumberFormatXml(_nameSpaceManager, n);
                    NumberFormats.Add(nf.Id, nf);
                    if (nf.NumFmtId >= NumberFormats.NextId) NumberFormats.NextId = nf.NumFmtId + 1;
                }
            }

            //Fonts
            XmlNode fontNode = GetNode(FontsPath);
            foreach (XmlNode n in fontNode)
            {
                ExcelFontXml f = new ExcelFontXml(_nameSpaceManager, n);
                Fonts.Add(f.Id, f);
            }

            //Fills
            XmlNode fillNode = GetNode(FillsPath);
            foreach (XmlNode n in fillNode)
            {
                ExcelFillXml f;
                if (n.FirstChild != null && n.FirstChild.LocalName == "gradientFill")
                {
                    f = new ExcelGradientFillXml(_nameSpaceManager, n);
                }
                else
                {
                    f = new ExcelFillXml(_nameSpaceManager, n);
                }
                Fills.Add(f.Id, f);
            }

            //Borders
            XmlNode borderNode = GetNode(BordersPath);
            foreach (XmlNode n in borderNode)
            {
                ExcelBorderXml b = new ExcelBorderXml(_nameSpaceManager, n);
                Borders.Add(b.Id, b);
            }

            //cellStyleXfs
            XmlNode styleXfsNode = GetNode(CellStyleXfsPath);
            if (styleXfsNode != null)
            {
                foreach (XmlNode n in styleXfsNode)
                {
                    ExcelXfs item = new ExcelXfs(_nameSpaceManager, n, this);
                    CellStyleXfs.Add(item.Id, item);
                }
            }

            XmlNode styleNode = GetNode(CellXfsPath);
            for (int i = 0; i < styleNode.ChildNodes.Count; i++)
            {
                XmlNode n = styleNode.ChildNodes[i];
                ExcelXfs item = new ExcelXfs(_nameSpaceManager, n, this);
                CellXfs.Add(item.Id, item);
            }

            //cellStyle
            XmlNode namedStyleNode = GetNode(CellStylesPath);
            if (namedStyleNode != null)
            {
                foreach (XmlNode n in namedStyleNode)
                {
                    ExcelNamedStyleXml item = new ExcelNamedStyleXml(_nameSpaceManager, n, this);
                    if(item.BuildInId==0)
                    {
                        _normalStyle = item;
                    }
                    NamedStyles.Add(item.Name, item);
                }
            }

            DxfStyleHandler.Load(_wb, this, Dxfs, DxfsPath);
            LoadTableStyles();
            LoadSlicerStyles();
        }

        private void LoadSlicerStyles()
        {
            //Slicer Styles
            XmlNode slicerStylesNode = GetNode(SlicerStylesPath);
            if (slicerStylesNode != null)
            {
                DxfStyleHandler.Load(_wb, this, DxfsSlicers, DxfSlicerStylesPath);    //Slicer styles have their own dxf collection inside the extLst.
                foreach (XmlNode n in slicerStylesNode)
                {
                    var name = n.Attributes["name"]?.Value;
                    XmlNode tableStyleNode;
                    if (_slicerTableStyleNodes.ContainsKey(name))
                    {
                        tableStyleNode = _slicerTableStyleNodes[name];
                    }
                    else if(TableStyles._dic.ContainsKey(name))
                    {
                        tableStyleNode = TableStyles[name].TopNode;
                    }
                    else
                    {
                        tableStyleNode = null;
                    }
                    var item = new ExcelSlicerNamedStyle(_nameSpaceManager, n, tableStyleNode, this);
                    SlicerStyles.Add(item.Name, item);
                }

            }
        }

        private void LoadTableStyles()
        {
            //Table Styles
            XmlNode tableStyleNode = GetNode(TableStylesPath);
            if (tableStyleNode != null)
            {
                foreach (XmlNode n in tableStyleNode)
                {
                    ExcelTableNamedStyleBase item;
                    var pivot = !(n.Attributes["pivot"]?.Value == "0");
                    var table = !(n.Attributes["table"]?.Value == "0");
                    if (pivot || table)
                    {
                        if (pivot==false)
                        {
                            item = new ExcelTableNamedStyle(_nameSpaceManager, n, this);
                        }
                        else if (table==false)
                        {
                            item = new ExcelPivotTableNamedStyle(_nameSpaceManager, n, this);
                        }
                        else
                        {
                            item = new ExcelTableAndPivotTableNamedStyle(_nameSpaceManager, n, this);
                        }
                        TableStyles.Add(item.Name, item);
                    }
                    else
                    {
                        //Styles for slicers and timelines. Timelines are currently unsupported.
                        var name = n.Attributes["name"]?.Value;
                        if (string.IsNullOrEmpty(name) == false)
                        {
                            _slicerTableStyleNodes.Add(name, n);
                        }
                    }
                }
            }
        }

        internal ExcelNamedStyleXml GetNormalStyle()
        {
            if (_normalStyle == null)
            {
                foreach (var style in NamedStyles)
                {
                    if (style.BuildInId == 0)
                    {
                        _normalStyle = style;
                        break;
                    }
                }
                if (_normalStyle==null && _wb.Styles.NamedStyles.Count > 0)
                {
                    return _wb.Styles.NamedStyles[0];
                }
            }
            return _normalStyle;
        }

        internal ExcelStyle GetStyleObject(int Id, int PositionID, string Address)
        {
            if (Id < 0) Id = 0;
            return new ExcelStyle(this, PropertyChange, PositionID, Address, Id);
        }
        /// <summary>
        /// Handels changes of properties on the style objects
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        internal int PropertyChange(StyleBase sender, Style.StyleChangeEventArgs e)
        {
            var address = new ExcelAddressBase(e.Address);
            var ws = _wb.Worksheets[e.PositionID];
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            //Set single address
            lock (ws._values)
            {
                if (address.Addresses == null)
                {
                    SetStyleAddress(sender, e, address, ws, ref styleCashe);
                }
                else
                {
                    //Handle multiaddresses
                    foreach (var innerAddress in address.Addresses)
                    {
                        SetStyleAddress(sender, e, innerAddress, ws, ref styleCashe);
                    }
                }
            }
            return 0;
        }
        private void SetStyleAddress(StyleBase sender, Style.StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, ref Dictionary<int, int> styleCashe)
        {
            if (address.Start.Column == 0 || address.Start.Row == 0)
            {
                throw (new Exception("error address"));
            }
            //Columns
            else if (address.Start.Row == 1 && address.End.Row == ExcelPackage.MaxRows)
            {
                SetStyleFullColumn(sender, e, address, ws, styleCashe);
            }
            //Rows
            else if (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns)
            {
                SetStyleFullRow(sender, e, address, ws, styleCashe);
            }
            //Cellrange
            else
            {
                SetStyleCells(sender, e, address, ws, styleCashe);
            }
        }

        private void SetStyleCells(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
        {
            ws._values.EnsureColumnsExists(address._fromCol, address._toCol);
            var rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
            var colCache = new Dictionary<int, ExcelValue>(address.End.Column - address.Start.Column + 1);
            var cellEnum = new CellStoreEnumerator<ExcelValue>(ws._values, address.Start.Row, address.Start.Column, address.End.Row, address.End.Column);
            var hasEnumValue = cellEnum.Next();
            for (int row = address._fromRow; row <= address._toRow; row++)
            {
                for (int col = address._fromCol; col <= address._toCol; col++)
                {
                    ExcelValue value;
                    if (hasEnumValue && row == cellEnum.Row && col == cellEnum.Column)
                    {
                        value = cellEnum.Value;
                        hasEnumValue = cellEnum.Next();
                    }
                    else
                    {
                        value = new ExcelValue { _styleId = 0 };
                    }
                    var s = value._styleId;
                    if (s == 0)
                    {
                        // get row styleId with cache
                        if (rowCache.ContainsKey(row))
                        {
                            s = rowCache[row];
                        }
                        else
                        {
                            s = ws._values.GetValue(row, 0)._styleId;
                            rowCache.Add(row, s);
                        }
                        if (s == 0)
                        {
                            // get column styleId with cache
                            if (colCache.ContainsKey(col))
                            {
                                s = colCache[col]._styleId;
                            }
                            else
                            {
                                var v = ws._values.GetValue(0, col);
                                if (v._value == null)
                                {
                                    if(colCache.TryGetValue(col, out ExcelValue ev))
                                    {
                                        s = ev._styleId;
                                    }
                                    else
                                    {
                                        int r = 0, c = col;
                                        if (ws._values.PrevCell(ref r, ref c))
                                        {
                                            if (!colCache.ContainsKey(c)) colCache.Add(c, ws._values.GetValue(0, c));
                                            var val = colCache[c];
                                            var colObj = val._value as ExcelColumn;
                                            if (colObj != null && colObj.ColumnMax >= col) //Fixes issue 15174
                                            {
                                                s = val._styleId;
                                            }
                                        }
                                        else
                                        {
                                            colCache.Add(col, new ExcelValue() { _styleId = 0 });
                                        }
                                    }
                                }
                                else
                                {
                                    colCache.Add(col, v);
                                    s = v._styleId;
                                }
                            }
                        }
                    }
                    if (styleCashe.ContainsKey(s))
                    {
                        ws._values.SetValue(row, col, new ExcelValue { _value = value._value, _styleId = styleCashe[s] });
                    }
                    else
                    {
                        ExcelXfs st;
                        if (s==0)
                        {
                            var ns=GetNormalStyle();   //Get the xfs id for the normal style.
                            if (ns==null || ns.StyleXfId<0)
                            {
                                st = CellXfs[0];
                            }
                            else
                            {
                                st = CellStyleXfs[ns.StyleXfId];
                            }

                        }
                        else
                        {
                            st = CellXfs[s];
                        }

                        int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(s, newId);
                        ws._values.SetValue(row, col, new ExcelValue { _value = value._value, _styleId = newId });
                    }
                }
            }
        }

        private bool GetFromCache(Dictionary<int, ExcelValue> colCache, int col, ref int s)
        {
            var c = col;
            while (!colCache.ContainsKey(--c))
            {
                if (c <= 0) return false;
            }
            var colObj = (ExcelColumn)(colCache[c]._value);
            if (colObj != null && colObj.ColumnMax >= col) //Fixes issue 15174
            {
                s = colCache[c]._styleId;
            }
            else
            {
                s = 0;
            }
            return true;
        }

        private void SetStyleFullRow(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
        {
            for (int rowNum = address.Start.Row; rowNum <= address.End.Row; rowNum++)
            {
                var s = ws.GetStyleInner(rowNum, 0);
                if (s == 0)
                {
                    //iterate all columns and set the row to the style of the last column
                    var cse = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
                    var cs = 0;
                    while (cse.Next())
                    {
                        cs = cse.Value._styleId;
                        if (cs == 0) continue;
                        var c = ws.GetValueInner(cse.Row, cse.Column) as ExcelColumn;
                        if (c != null && c.ColumnMax < ExcelPackage.MaxColumns)
                        {
                            for (int col = c.ColumnMin; col < c.ColumnMax; col++)
                            {
                                if (!ws.ExistsStyleInner(rowNum, col))
                                {
                                    ws.SetStyleInner(rowNum, col, cs);
                                }
                            }
                        }
                    }
                    ws.SetStyleInner(rowNum, 0, cs);
                    cse.Dispose();
                }
                if (styleCashe.ContainsKey(s))
                {
                    ws.SetStyleInner(rowNum, 0, styleCashe[s]);
                }
                else
                {
                    ExcelXfs st = CellXfs[s];
                    int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                    styleCashe.Add(s, newId);
                    ws.SetStyleInner(rowNum, 0, newId);
                }
            }

            //Update individual cells 
            var cse2 = new CellStoreEnumerator<ExcelValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
            while (cse2.Next())
            {
                var s = cse2.Value._styleId;
                if (s == 0) continue;
                if (styleCashe.ContainsKey(s))
                {
                    ws.SetStyleInner(cse2.Row, cse2.Column, styleCashe[s]);
                }
                else
                {
                    ExcelXfs st = CellXfs[s];
                    int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                    styleCashe.Add(s, newId);
                    ws.SetStyleInner(cse2.Row, cse2.Column, newId);
                }
            }

            //Update cells with styled rows
            cse2 = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, address._toCol);
            while (cse2.Next())
            {
                if (cse2.Value._styleId == 0) continue;
                for (int r = address._fromRow; r <= address._toRow; r++)
                {
                    if (!ws.ExistsStyleInner(r, cse2.Column))
                    {
                        var s = cse2.Value._styleId;
                        if (styleCashe.ContainsKey(s))
                        {
                            ws.SetStyleInner(r, cse2.Column, styleCashe[s]);
                        }
                        else
                        {
                            ExcelXfs st = CellXfs[s];
                            int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                            styleCashe.Add(s, newId);
                            ws.SetStyleInner(r, cse2.Column, newId);
                        }
                    }
                }
            }
        }

        private void SetStyleFullColumn(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
        {
            ExcelColumn column;
            int col = address.Start.Column, row = 0;
            bool isNew;
            //Get the startcolumn
            object o = null;
            if (!ws.ExistsValueInner(0, address.Start.Column, ref o))
            {
                column = ws.Column(address.Start.Column);
                isNew = true;
            }
            else
            {
                column = (ExcelColumn)o;
                isNew = false;
            }
            var prevColumMax = column.ColumnMax;
            while (column.ColumnMin <= address.End.Column)
            {
                if (column.ColumnMin > prevColumMax + 1)
                {
                    var newColumn = ws.Column(prevColumMax + 1);
                    newColumn.ColumnMax = column.ColumnMin - 1;
                    AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
                }
                if (column.ColumnMax > address.End.Column)
                {
                    var newCol = ws.CopyColumn(column, address.End.Column + 1, column.ColumnMax);
                    column.ColumnMax = address.End.Column;
                }
                var s = ws.GetStyleInner(0, column.ColumnMin);
                AddNewStyleColumn(sender, e, ws, styleCashe, column, s);

                //index++;
                prevColumMax = column.ColumnMax;
                if (!ws._values.NextCell(ref row, ref col) || row > 0)
                {
                    if (column._columnMax == address.End.Column)
                    {
                        break;
                    }

                    if (isNew)
                    {
                        column._columnMax = address.End.Column;
                    }
                    else
                    {
                        var newColumn = ws.Column(column._columnMax + 1);
                        newColumn.ColumnMax = address.End.Column;
                        AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
                        column = newColumn;
                    }
                    break;
                }
                else
                {
                    column = (ws.GetValueInner(0, col) as ExcelColumn);
                }
            }

            if (column._columnMax < address.End.Column)
            {
                var newCol = ws.Column(column._columnMax + 1) as ExcelColumn;
                newCol._columnMax = address.End.Column;

                var s = ws.GetStyleInner(0, column.ColumnMin);
                if (styleCashe.ContainsKey(s))
                {
                    ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
                }
                else
                {
                    ExcelXfs st = CellXfs[s];
                    int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                    styleCashe.Add(s, newId);
                    ws.SetStyleInner(0, column.ColumnMin, newId);
                }

                column._columnMax = address.End.Column;
            }

            //Set for individual cells in the span. We loop all cells here since the cells are sorted with columns first.
            var cse = new CellStoreEnumerator<ExcelValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);
            while (cse.Next())
            {
                if (cse.Column >= address.Start.Column &&
                    cse.Column <= address.End.Column &&
                    cse.Value._styleId != 0)
                {
                    if (styleCashe.ContainsKey(cse.Value._styleId))
                    {
                        ws.SetStyleInner(cse.Row, cse.Column, styleCashe[cse.Value._styleId]);
                    }
                    else
                    {
                        ExcelXfs st = CellXfs[cse.Value._styleId];
                        int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(cse.Value._styleId, newId);
                        ws.SetStyleInner(cse.Row, cse.Column, newId);
                    }
                }
            }

            if (!(address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns))
            {
                //Update cells with styled columns
                cse = new CellStoreEnumerator<ExcelValue>(ws._values, 1, 0, address._toRow, 0);
                while (cse.Next())
                {
                    if (cse.Value._styleId == 0) continue;
                    for (int c = address._fromCol; c <= address._toCol; c++)
                    {
                        if (!ws.ExistsStyleInner(cse.Row, c))
                        {
                            if (styleCashe.ContainsKey(cse.Value._styleId))
                            {
                                ws.SetStyleInner(cse.Row, c, styleCashe[cse.Value._styleId]);
                            }
                            else
                            {
                                ExcelXfs st = CellXfs[cse.Value._styleId];
                                int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                                styleCashe.Add(cse.Value._styleId, newId);
                                ws.SetStyleInner(cse.Row, c, newId);
                            }
                        }
                    }
                }
            }
        }

        private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s)
        {
            if (styleCashe.ContainsKey(s))
            {
                ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
            }
            else
            {
                ExcelXfs st = CellXfs[s];
                int newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                styleCashe.Add(s, newId);
                ws.SetStyleInner(0, column.ColumnMin, newId);
            }
        }
        internal int GetStyleId(ExcelWorksheet ws, int row, int col)
        {
            int v = 0;
            if (ws.ExistsStyleInner(row, col, ref v))
            {
                return v;
            }
            else
            {
                if (ws.ExistsStyleInner(row, 0, ref v)) //First Row
                {
                    return v;
                }
                else // then column
                {
                    if (ws.ExistsStyleInner(0, col, ref v))
                    {
                        return v;
                    }
                    else
                    {
                        int r = 0, c = col;
                        if (ws._values.PrevCell(ref r, ref c))
                        {
                            //var column=ws.GetValueInner(0,c) as ExcelColumn;
                            var val = ws._values.GetValue(0, c);
                            var column = (ExcelColumn)(val._value);
                            if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                            {
                                //return ws.GetStyleInner(0, c);
                                return val._styleId;
                            }
                            else
                            {
                                return 0;
                            }
                        }
                        else
                        {
                            return 0;
                        }
                    }

                }
            }

        }
        /// <summary>
        /// Handles property changes on Named styles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        internal int NamedStylePropertyChange(StyleBase sender, Style.StyleChangeEventArgs e)
        {

            int index = NamedStyles.FindIndexById(e.Address);
            if (index >= 0)
            {
                if(e.StyleClass == eStyleClass.Font && (e.StyleProperty==eStyleProperty.Name || e.StyleProperty == eStyleProperty.Size) && NamedStyles[index].BuildInId==0)
                {
                    foreach(var ws in _wb.Worksheets)
                    {
                        ws.NormalStyleChange();
                    }
                }
                int newId = CellStyleXfs[NamedStyles[index].StyleXfId].GetNewID(CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                int prevIx = NamedStyles[index].StyleXfId;
                NamedStyles[index].StyleXfId = newId;
                NamedStyles[index].Style.Index = newId;

                NamedStyles[index].XfId = int.MinValue;
                foreach (var style in CellXfs)
                {
                    if (style.XfId == prevIx)
                    {
                        style.XfId = newId;
                    }
                }
            }
            return 0;
        }
        /// <summary>
        /// Contains all numberformats for the package
        /// </summary>
        public ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats = new ExcelStyleCollection<ExcelNumberFormatXml>();
        /// <summary>
        /// Contains all font styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelFontXml> Fonts = new ExcelStyleCollection<ExcelFontXml>();
        /// <summary>
        /// Contains all fill styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelFillXml> Fills = new ExcelStyleCollection<ExcelFillXml>();
        /// <summary>
        /// Contain all border styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelBorderXml> Borders = new ExcelStyleCollection<ExcelBorderXml>();
        /// <summary>
        /// Contain all named cell styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelXfs> CellStyleXfs = new ExcelStyleCollection<ExcelXfs>();
        /// <summary>
        /// Contain all cell styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelXfs> CellXfs = new ExcelStyleCollection<ExcelXfs>();
        /// <summary>
        /// Contain all named styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = new ExcelStyleCollection<ExcelNamedStyleXml>();
        /// <summary>
        /// Contain all table styles for the package. Tables styles can be used to customly format tables and pivot tables.
        /// </summary>
        public ExcelNamedStyleCollection<ExcelTableNamedStyleBase> TableStyles = new ExcelNamedStyleCollection<ExcelTableNamedStyleBase>();
        /// <summary>
        /// Contain all slicer styles for the package. Tables styles can be used to customly format tables and pivot tables.
        /// </summary>
        public ExcelNamedStyleCollection<ExcelSlicerNamedStyle> SlicerStyles = new ExcelNamedStyleCollection<ExcelSlicerNamedStyle>();
        /// <summary>
        /// Contain differential formatting styles for the package. This collection does not contain style records for slicers.
        /// </summary>
        public ExcelStyleCollection<ExcelDxfStyleBase> Dxfs = new ExcelStyleCollection<ExcelDxfStyleBase>();
        internal ExcelStyleCollection<ExcelDxfStyleBase> DxfsSlicers = new ExcelStyleCollection<ExcelDxfStyleBase>();
        internal Dictionary<string, XmlNode> _slicerTableStyleNodes = new Dictionary<string, XmlNode>(StringComparer.InvariantCultureIgnoreCase);
        internal string Id
        {
            get { return ""; }
        }
        /// <summary>
        /// Creates a named style that can be applied to cells in the worksheet.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <returns>A named style object that can be custumized</returns>
        public ExcelNamedStyleXml CreateNamedStyle(string name)
        {
            return CreateNamedStyle(name, null);
        }
        /// <summary>
        /// Creates a named style that can be applied to cells in the worksheet.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="Template">A template style</param>
        /// <returns>A named style object that can be custumized</returns>
        public ExcelNamedStyleXml CreateNamedStyle(string name, ExcelStyle Template)
        {
            if (_wb.Styles.NamedStyles.ExistsKey(name))
            {
                throw new Exception(string.Format("Key {0} already exists in collection", name));
            }

            ExcelNamedStyleXml style;
            style = new ExcelNamedStyleXml(NameSpaceManager, this);
            int xfIdCopy, positionID;
            ExcelStyles styles;
            bool isTemplateNamedStyle;
            if (Template == null)
            {
                var ns = _wb.Styles.GetNormalStyle();
                if (ns != null)
                {
                    xfIdCopy = ns.StyleXfId;
                }
                else
                {
                    xfIdCopy = -1;
                }
                positionID = -1;
                styles = this;
                isTemplateNamedStyle = true;
            }
            else
            {
                isTemplateNamedStyle = Template.PositionID == -1;
                if (Template.PositionID < 0 && Template.Styles==this)
                {
                    xfIdCopy = Template.Index;
                    
                    positionID=Template.PositionID;
                    styles = this;
                }
                else
                {
                    xfIdCopy = Template.Index;
                    positionID = -1;
                    styles = Template.Styles;
                }
            }
            //Clone namedstyle
            if (xfIdCopy >= 0)
            {
                int styleXfId = CloneStyle(styles, xfIdCopy, true, false, isTemplateNamedStyle);
                CellStyleXfs[styleXfId].XfId = CellStyleXfs.Count - 1;
                style.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, styleXfId);
                style.StyleXfId = styleXfId;
            }
            else
            {
                style.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, 0);
                style.StyleXfId = 0;
            }

            style.Name = name;
            int ix =_wb.Styles.NamedStyles.Add(style.Name, style);
            style.Style.SetIndex(ix);
            return style;
        }
        /// <summary>
        /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <returns>The table style object</returns>
        public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name)
        {
            ValidateTableStyleName(name);
            var node = (XmlElement)CreateNode("d:tableStyles/d:tableStyle", false, true);
            node.SetAttribute("table", "0");
            var s = new ExcelPivotTableNamedStyle(NameSpaceManager, node, this)
            {
                Name = name
            };
            TableStyles.Add(name, s);
            return s;
        }
        /// <summary>
        /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name, PivotTableStyles templateStyle)
        {
            var s = CreatePivotTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        /// <summary>
        /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
        {
            var s = CreatePivotTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }

        /// <summary>
        /// Creates a tables style only visible for tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <returns>The table style object</returns>
        public ExcelTableNamedStyle CreateTableStyle(string name)
        {
            ValidateTableStyleName(name);
            var node = (XmlElement)CreateNode("d:tableStyles/d:tableStyle", false, true);
            node.SetAttribute("pivot", "0");
            var s = new ExcelTableNamedStyle(NameSpaceManager, node, this)
            {
                Name = name
            };
            TableStyles.Add(name, s);
            return s;
        }
        /// <summary>
        /// Creates a tables style only visible for tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelTableNamedStyle CreateTableStyle(string name, TableStyles templateStyle)
        {
            var s = CreateTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        /// <summary>
        /// Creates a tables style only visible for tables and with elements specific to pivot tables.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelTableNamedStyle CreateTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
        {
            var s = CreateTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }

        /// <summary>
        /// Creates a tables visible for tables and pivot tables and with elements for both.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <returns>The table style object</returns>
        public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name)
        {
            ValidateTableStyleName(name);
            var node = (XmlElement)CreateNode("d:tableStyles/d:tableStyle", false, true);
            var s = new ExcelTableAndPivotTableNamedStyle(NameSpaceManager, node, this)
            {
                Name = name
            };
            TableStyles.Add(name, s);
            return s;
        }
        /// <summary>
        /// Creates a tables visible for tables and pivot tables and with elements for both.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, TableStyles templateStyle)
        {
            if (templateStyle == Table.TableStyles.Custom)
            {
                throw new ArgumentException("Cant use template style Custom. To use a custom style, please use the ´PivotTableStyles´ overload of this method.", nameof(templateStyle));
            }

            var s = CreateTableAndPivotTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        /// <summary>
        /// Creates a tables visible for tables and pivot tables and with elements for both.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The built-in pivot table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, PivotTableStyles templateStyle)
        {
            if (templateStyle == PivotTableStyles.Custom)
            {
                throw new ArgumentException("Cant use template style Custom. To use a custom style, please use the ´ExcelTableNamedStyleBase´ overload of this method.", nameof(templateStyle));
            }

            var s = CreateTableAndPivotTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        /// <summary>
        /// Creates a tables visible for tables and pivot tables and with elements for both.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The table style to use as a template for this custom style</param>
        /// <returns>The table style object</returns>
        public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
        {
            var s = CreateTableAndPivotTableStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }

        /// <summary>
        /// Creates a custom slicer style.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <returns>The slicer style object</returns>
        public ExcelSlicerNamedStyle CreateSlicerStyle(string name)
        {
            ValidateTableStyleName(name);

            //Create the matching table style
            var tableStyleNode = (XmlElement)CreateNode("d:tableStyles/d:tableStyle", false, true);
            tableStyleNode.SetAttribute("table", "0");
            tableStyleNode.SetAttribute("pivot", "0");
            tableStyleNode.SetAttribute("name", name);
            _slicerTableStyleNodes.Add(name, tableStyleNode);

            //The dxfs collection must be created before the slicer styles collection
            GetOrCreateExtLstSubNode(ExtLstUris.SlicerStylesDxfCollectionUri, "x14");

            var extNode = GetOrCreateExtLstSubNode(ExtLstUris.SlicerStylesUri, "x14");
            var extHelper = XmlHelperFactory.Create(NameSpaceManager, extNode);
            if (extNode.ChildNodes.Count==0)
            {
                var slicersNode=(XmlElement)extHelper.CreateNode("x14:slicerStyles", false, true);
                slicersNode.SetAttribute("defaultSlicerStyle", "SlicerStyleLight1");    //defaultSlicerStyle is required
                extHelper.TopNode = slicersNode;
            }
            else
            {
                extHelper.TopNode = extNode.FirstChild;
            }
            var node = (XmlElement)extHelper.CreateNode("x14:slicerStyle", false, true);

             var s = new ExcelSlicerNamedStyle(NameSpaceManager, node, tableStyleNode, this)
            {
                Name = name
            };
            SlicerStyles.Add(name, s);
            return s;
        }
        /// <summary>
        /// Creates a custom slicer style.
        /// </summary>
        /// <param name="name">The name of the style</param>
        /// <param name="templateStyle">The slicer style to use as a template for this custom style</param>
        /// <returns>The slicer style object</returns>
        public ExcelSlicerNamedStyle CreateSlicerStyle(string name, eSlicerStyle templateStyle)
        {
            if(templateStyle==eSlicerStyle.Custom)
            {
                throw new ArgumentException("Cant use template style Custom. To use a custom style, please use the ´ExcelSlicerNamedStyle´ overload of this method.", nameof(templateStyle));
            }
            var s = CreateSlicerStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        HashSet<string> tableStyleNames=new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private void ValidateTableStyleName(string name)
        {
            if(tableStyleNames.Count==0)
            {
                Enum.GetNames(typeof(TableStyles)).Select(x => tableStyleNames.Add("TableStyle"+x));
                Enum.GetNames(typeof(PivotTableStyles)).Select(x => tableStyleNames.Add("PivotTableStyle" + x));
                Enum.GetNames(typeof(eSlicerStyle)).Select(x => tableStyleNames.Add("SlicerStyle" + x));
            }
            if(tableStyleNames.Contains(name) || TableStyles.ExistsKey(name) || SlicerStyles.ExistsKey(name))
            {
                throw (new ArgumentException($"Table style name is not unique : {name}", "name"));
            }
        }
        /// <summary>
        /// Creates a custom named slicer style from another style.
        /// </summary>
        /// <param name="name">The name of the style.</param>
        /// <param name="templateStyle">The slicer style to us as template.</param>
        /// <returns></returns>
        public ExcelSlicerNamedStyle CreateSlicerStyle(string name, ExcelSlicerNamedStyle templateStyle)
        {
            var s = CreateSlicerStyle(name);
            s.SetFromTemplate(templateStyle);
            return s;
        }
        /// <summary>
        /// Update the changes to the Style.Xml file inside the package.
        /// This will remove any unused styles from the collections.
        /// </summary>
        public void UpdateXml()
        {
            RemoveUnusedStyles();

            int normalIx = GetNormalStyleIndex();

            UpdateNumberFormatXml(normalIx);
            UpdateFontXml(normalIx);
            UpdateFillXml();
            UpdateBorderXml();
            UpdateNamedStylesAndXfs(normalIx);

            DxfStyleHandler.UpdateDxfXml(_wb);
        }

        private void UpdateNamedStylesAndXfs(int normalIx)
        {
            //Create the cellStyleXfs element            
            XmlNode styleXfsNode = GetNode(CellStyleXfsPath);
            if (styleXfsNode == null)
            {
                if (CellStyleXfs.Count > 0)
                {
                    styleXfsNode = CreateNode(CellStyleXfsPath);
                }
            }
            else
            {
                styleXfsNode?.RemoveAll();
            }
            //NamedStyles
            int count = normalIx > -1 ? 1 : 0;  //If we have a normal style, we make sure it's added first.

            XmlNode cellStyleNode = GetNode(CellStylesPath);
            if (cellStyleNode == null)
            {
                if (NamedStyles.Count > 0)
                {
                    cellStyleNode  = CreateNode(CellStylesPath);
                }
            }
            else
            {
                cellStyleNode.RemoveAll();
            }
            
            XmlNode cellXfsNode = GetNode(CellXfsPath);
            cellXfsNode.RemoveAll();
            int xfsCount = 0;
            if (CellStyleXfs.Count > 0)
            {
                if (normalIx >= 0)
                {
                    NamedStyles[normalIx].newID = 0;
                    AddNamedStyle(0, styleXfsNode, cellXfsNode, NamedStyles[normalIx]);
                    cellXfsNode.AppendChild(CellStyleXfs[NamedStyles[normalIx].StyleXfId].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                }
                else
                {
                    styleXfsNode.AppendChild(CellStyleXfs[0].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
                    CellStyleXfs[0].newID = 0;
                    cellXfsNode.AppendChild(CellStyleXfs[0].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                }
                xfsCount++;
            }

            foreach (ExcelNamedStyleXml style in NamedStyles)
            {
                if (style.BuildInId != 0)
                {
                    AddNamedStyle(count++, styleXfsNode, cellXfsNode, style);
                }
                else
                {
                    style.newID = 0;
                }
                cellStyleNode.AppendChild(style.CreateXmlNode(_styleXml.CreateElement("cellStyle", ExcelPackage.schemaMain)));
            }

            if (cellStyleNode != null)
            {
                var cellStyleElement = (cellStyleNode as XmlElement);
                cellStyleElement.SetAttribute("count", cellStyleElement.ChildNodes.Count.ToString(CultureInfo.InvariantCulture));
            }
            if (styleXfsNode != null)
            {
                var styleXfsElement = (styleXfsNode as XmlElement);
                styleXfsElement.SetAttribute("count", styleXfsElement.ChildNodes.Count.ToString(CultureInfo.InvariantCulture));
            }

            //CellStyle
            int xfix = 0;
            foreach (ExcelXfs xf in CellXfs)
            {
                if (xf.useCnt > 0 && !(xfix==0 && normalIx >= 0))
                {
                    cellXfsNode.AppendChild(xf.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                    xf.newID = xfsCount;
                    xfsCount++;
                }
                xfix++;
            }
            (cellXfsNode as XmlElement).SetAttribute("count", xfsCount.ToString(CultureInfo.InvariantCulture));

        }

        private void UpdateBorderXml()
        {
            //Borders
            int count = 0;
            XmlNode bordersNode = GetNode(BordersPath);
            bordersNode.RemoveAll();
            Borders[0].useCnt = 1;    //Must exist blank;
            foreach (ExcelBorderXml border in Borders)
            {
                if (border.useCnt > 0)
                {
                    bordersNode.AppendChild(border.CreateXmlNode(_styleXml.CreateElement("border", ExcelPackage.schemaMain)));
                    border.newID = count;
                    count++;
                }
            }
            (bordersNode as XmlElement).SetAttribute("count", count.ToString());
        }

        private int UpdateFillXml()
        {
            //Fills
            int count = 0;
            XmlNode fillsNode = GetNode(FillsPath);
            fillsNode.RemoveAll();
            Fills[0].useCnt = 1;    //Must exist (none);  
            Fills[1].useCnt = 1;    //Must exist (gray125);
            foreach (ExcelFillXml fill in Fills)
            {
                if (fill.useCnt > 0)
                {
                    fillsNode.AppendChild(fill.CreateXmlNode(_styleXml.CreateElement("fill", ExcelPackage.schemaMain)));
                    fill.newID = count;
                    count++;
                }
            }

            (fillsNode as XmlElement).SetAttribute("count", count.ToString());
            return count;
        }

        internal int GetNormalStyleIndex()
        {
            int normalIx = NamedStyles.FindIndexByBuildInId(0);

            if (normalIx < 0)
            {
                normalIx = NamedStyles.FindIndexById("normal");
            }
            
           return normalIx;
        }

        private void UpdateFontXml(int normalIx)
        {
            //Font
            int count = 0;
            XmlNode fntNode = GetNode(FontsPath);
            fntNode.RemoveAll();
            int nfIx = -1;
            //Normal should be first in the collection
            if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Font.Index >= 0)
            {
                nfIx = NamedStyles[normalIx].Style.Font.Index;
                ExcelFontXml fnt = Fonts[nfIx];
                fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                fnt.newID = count++;
            }

            int ix = 0;
            foreach (ExcelFontXml fnt in Fonts)
            {
                if (fnt.useCnt > 0 && ix!=nfIx)
                {
                    fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                    fnt.newID = count;
                    count++;
                }
                ix++;
            }
            (fntNode as XmlElement).SetAttribute("count", count.ToString());
        }

        private void UpdateNumberFormatXml(int normalIx)
        {
            //NumberFormat
            XmlNode nfNode = GetNode(NumberFormatsPath);
            if (nfNode == null)
            {
                nfNode = CreateNode(NumberFormatsPath, true);
            }
            else
            {
                nfNode.RemoveAll();
            }

            int count = 0;
            if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Numberformat.NumFmtID >= 164)
            {
                ExcelNumberFormatXml nf = NumberFormats[NumberFormats.FindIndexById(NamedStyles[normalIx].Style.Numberformat.Id)];
                nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                nf.newID = count++;
            }

            //Add pivot Table formatting.
            foreach (var ws in _wb.Worksheets)
            {
                if (!(ws is ExcelChartsheet) && ws.HasLoadedPivotTables)
                {
                    foreach (var pt in ws.PivotTables)
                    {
                        foreach (var f in pt.Fields)
                        {
                            f.NumFmtId = GetNumFormatId(f.Format);
                            f.Cache.NumFmtId = GetNumFormatId(f.Cache.Format);
                        }
                        foreach (var df in pt.DataFields)
                        {
                            if (df.NumFmtId < 0 && df.Field.NumFmtId.HasValue)
                            {
                                df.NumFmtId = df.Field.NumFmtId.Value;
                            }
                        }
                    }
                }
            }

            foreach (ExcelNumberFormatXml nf in NumberFormats)
            {
                if (!nf.BuildIn /*&& nf.newID<0*/) //Buildin formats are not updated.
                {
                    nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                    nf.newID = count;
                    count++;
                }
            }

            (nfNode as XmlElement).SetAttribute("count", count.ToString());
        }

        private int? GetNumFormatId(string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                return null;
            }
            else
            {
                ExcelNumberFormatXml nf = null;
                if (NumberFormats.FindById(format, ref nf))
                {
                    return nf.NumFmtId;
                }
                else
                {
                    var id = NumberFormats.NextId++;
                    var item = new ExcelNumberFormatXml(NameSpaceManager, false) 
                    { 
                        Format = format,
                        NumFmtId = id
                    };
                    NumberFormats.Add(format, item);
                    return id;
                }
            }
        }

        private void AddNamedStyle(int id, XmlNode styleXfsNode,XmlNode cellXfsNode, ExcelNamedStyleXml style)
        {
            var styleXfs = CellStyleXfs[style.StyleXfId];
            styleXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
            styleXfs.newID = id;
            styleXfs.XfId = style.StyleXfId;

            if (style.XfId >= 0)
                style.XfId = CellXfs[style.XfId].newID;
            else
                style.XfId = 0;
        }
        private void RemoveUnusedStyles()
        {
            CellXfs[0].useCnt = 1; //First item is allways used.
            foreach (ExcelWorksheet sheet in _wb.Worksheets)
            {
                if (sheet is ExcelChartsheet) continue;
                var cse = new CellStoreEnumerator<ExcelValue>(sheet._values);
                while(cse.Next())
                {
                    var v = cse.Value._styleId;
                    if (v >= 0)
                    {
                        CellXfs[v].useCnt++;
                    }
                }
            }
            foreach (ExcelNamedStyleXml ns in NamedStyles)
            {
                CellStyleXfs[ns.StyleXfId].useCnt++;
            }

            foreach (ExcelXfs xf in CellXfs)
            {
                if (xf.useCnt > 0)
                {
                    if (xf.FontId >= 0) Fonts[xf.FontId].useCnt++;
                    if (xf.FillId >= 0) Fills[xf.FillId].useCnt++;
                    if (xf.BorderId >= 0) Borders[xf.BorderId].useCnt++;
                }
            }
            foreach (ExcelXfs xf in CellStyleXfs)
            {
                if (xf.useCnt > 0)
                {
                    if (xf.FontId >= 0) Fonts[xf.FontId].useCnt++;
                    if (xf.FillId >= 0) Fills[xf.FillId].useCnt++;
                    if (xf.BorderId >= 0) Borders[xf.BorderId].useCnt++;                    
                }
            }
        }
        internal int GetStyleIdFromName(string Name)
        {
            int i = NamedStyles.FindIndexById(Name);
            if (i >= 0)
            {
                int id = NamedStyles[i].XfId;
                if (id < 0)
                {
                    int styleXfId=NamedStyles[i].StyleXfId;
                    ExcelXfs newStyle = CellStyleXfs[styleXfId].Copy();
                    newStyle.XfId = styleXfId;
                    id = CellXfs.FindIndexById(newStyle.Id);
                    if (id < 0)
                    {
                        id = CellXfs.Add(newStyle.Id, newStyle);
                    }
                    NamedStyles[i].XfId=id;
                }
                return id;
            }
            else
            {
                return 0;
                //throw(new Exception("Named style does not exist"));        	         
            }
        }
   #region XmlHelpFunctions
        private int GetXmlNodeInt(XmlNode node)
        {
            int i;
            if (int.TryParse(GetXmlNode(node), out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        private string GetXmlNode(XmlNode node)
        {
            if (node == null)
            {
                return "";
            }
            if (node.Value != null)
            {
                return node.Value;
            }
            else
            {
                return "";
            }
        }

        #endregion

        internal int CloneStyle(ExcelStyles style, int styleID)
        {
            return CloneStyle(style, styleID, false, false, false);
        }
        internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle)
        {
            return CloneStyle(style, styleID, isNamedStyle, false, isNamedStyle);
        }
        internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle, bool allwaysAddCellXfs, bool isCellStyleXfs)
        {
            ExcelXfs xfs;
            lock (style)
            {
                if (isCellStyleXfs)
                {
                    xfs = style.CellStyleXfs[styleID];
                }
                else
                {
                    xfs = style.CellXfs[styleID];
                }
                ExcelXfs newXfs = xfs.Copy(this);
                //Numberformat
                if (xfs.NumberFormatId > 0)
                {
                    //rake36: Two problems here...
                    //rake36:  1. the first time through when format stays equal to String.Empty, it adds a string.empty to the list of Number Formats
                    //rake36:  2. when adding a second sheet, if the numberformatid == 164, it finds the 164 added by previous sheets but was using the array index
                    //rake36:      for the numberformatid

                    string format = string.Empty;
                    foreach (var fmt in style.NumberFormats)
                    {
                        if (fmt.NumFmtId == xfs.NumberFormatId)
                        {
                            format = fmt.Format;
                            break;
                        }
                    }
                    //rake36: Don't add another format if it's blank
                    if (!String.IsNullOrEmpty(format))
                    {
                        int ix = NumberFormats.FindIndexById(format);
                        if (ix < 0)
                        {
                            var item = new ExcelNumberFormatXml(NameSpaceManager) { Format = format, NumFmtId = NumberFormats.NextId++ };
                            NumberFormats.Add(format, item);
                            //rake36: Use the just added format id
                            newXfs.NumberFormatId = item.NumFmtId;
                        }
                        else
                        {
                            //rake36: Use the format id defined by the index... not the index itself
                            newXfs.NumberFormatId = NumberFormats[ix].NumFmtId;
                        }
                    }
                }

                //Font
                if (xfs.FontId > -1)
                {
                    int ix = Fonts.FindIndexById(xfs.Font.Id);
                    if (ix < 0)
                    {
                        ExcelFontXml item = style.Fonts[xfs.FontId].Copy();
                        ix = Fonts.Add(xfs.Font.Id, item);
                    }
                    newXfs.FontId = ix;
                }

                //Border
                if (xfs.BorderId > -1)
                {
                    int ix = Borders.FindIndexById(xfs.Border.Id);
                    if (ix < 0)
                    {
                        ExcelBorderXml item = style.Borders[xfs.BorderId].Copy();
                        ix = Borders.Add(xfs.Border.Id, item);
                    }
                    newXfs.BorderId = ix;
                }

                //Fill
                if (xfs.FillId > -1)
                {
                    int ix = Fills.FindIndexById(xfs.Fill.Id);
                    if (ix < 0)
                    {
                        var item = style.Fills[xfs.FillId].Copy();
                        ix = Fills.Add(xfs.Fill.Id, item);
                    }
                    newXfs.FillId = ix;
                }

                //Named style reference
                if (xfs.XfId > 0)
                {
                    if(style._wb!=_wb && allwaysAddCellXfs==false) //Not the same workbook, copy the namedstyle to the workbook or match the id
                    {
                        var nsFind = style.NamedStyles.ToDictionary(d => (d.StyleXfId));
                        if (nsFind.ContainsKey(xfs.XfId))
                        {
                            var st = nsFind[xfs.XfId];
                            if (NamedStyles.ExistsKey(st.Name))
                            {
                                newXfs.XfId = NamedStyles.FindIndexById(st.Name);
                            }
                            else
                            {
                                var ns = CreateNamedStyle(st.Name, st.Style);
                                newXfs.XfId = NamedStyles.Count - 1;
                            }
                        }
                    }
                    else
                    {
                        var id = style.CellStyleXfs[xfs.XfId].Id;
                        var newId = CellStyleXfs.FindIndexById(id);
                        if (newId >= 0)
                        {
                            newXfs.XfId = newId;
                        }
                    }
                }

                int index;
                if (isNamedStyle && allwaysAddCellXfs==false)
                {
                    index = CellStyleXfs.Add(newXfs.Id, newXfs);
                }
                else
                {
                    if (allwaysAddCellXfs)
                    {
                        index = CellXfs.Add(newXfs.Id, newXfs);
                    }
                    else
                    {
                        index = CellXfs.FindIndexById(newXfs.Id);
                        if (index < 0)
                        {
                            index = CellXfs.Add(newXfs.Id, newXfs);
                        }
                    }
                }
                return index;
            }
        }

        internal ExcelDxfStyleLimitedFont GetDxfLimitedFont(int? dxfId)
        {
            if (dxfId.HasValue && dxfId < Dxfs.Count)
            {
                return Dxfs[dxfId.Value].ToDxfLimitedStyle();
            }
            else
            {
                return new ExcelDxfStyleLimitedFont(NameSpaceManager, null, this, null);
            }
        }
        internal ExcelDxfStyle GetDxf(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
        {
            if(dxfId.HasValue && dxfId < Dxfs.Count)
            {
                return Dxfs[dxfId.Value].ToDxfStyle();
            }
            else
            {
                return new ExcelDxfStyle(NameSpaceManager, null, this, callback);
            }
        }
        internal ExcelDxfSlicerStyle GetDxfSlicer(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
        {
            if (dxfId.HasValue && dxfId < Dxfs.Count)
            {
                return Dxfs[dxfId.Value].ToDxfSlicerStyle();
            }
            else
            {
                return new ExcelDxfSlicerStyle(NameSpaceManager, null, this, callback);
            }
        }
        internal ExcelDxfTableStyle GetDxfTable(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
        {
            if (dxfId.HasValue && dxfId < Dxfs.Count)
            {
                return Dxfs[dxfId.Value].ToDxfTableStyle();
            }
            else
            {
                return new ExcelDxfTableStyle(NameSpaceManager, null, this, callback);
            }
        }

        internal ExcelDxfSlicerStyle GetDxfSlicer(int? dxfId)
        {
            if (dxfId.HasValue && dxfId < DxfsSlicers.Count)
            {
                return DxfsSlicers[dxfId.Value].ToDxfSlicerStyle();
            }
            else
            {
                return new ExcelDxfSlicerStyle(NameSpaceManager, null, this, null);
            }
        }

    }
}
