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

namespace OfficeOpenXml
{
	/// <summary>
	/// Containts all shared cell styles for a workbook
	/// </summary>
    public sealed class ExcelStyles : XmlHelper
    {
        const string NumberFormatsPath = "d:styleSheet/d:numFmts";
        const string FontsPath = "d:styleSheet/d:fonts";
        const string FillsPath = "d:styleSheet/d:fills";
        const string BordersPath = "d:styleSheet/d:borders";
        const string CellStyleXfsPath = "d:styleSheet/d:cellStyleXfs";
        const string CellXfsPath = "d:styleSheet/d:cellXfs";
        const string CellStylesPath = "d:styleSheet/d:cellStyles";
        const string dxfsPath = "d:styleSheet/d:dxfs";

        //internal Dictionary<int, ExcelXfs> Styles = new Dictionary<int, ExcelXfs>();
        XmlDocument _styleXml;
        ExcelWorkbook _wb;
        XmlNamespaceManager _nameSpaceManager;
        internal int _nextDfxNumFmtID = 164;
        internal ExcelStyles(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml)
        {       
            _styleXml=xml;
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
            //NumberFormats
            ExcelNumberFormatXml.AddBuildIn(NameSpaceManager, NumberFormats);
            XmlNode numNode = _styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
            if (numNode != null)
            {
                foreach (XmlNode n in numNode)
                {
                    ExcelNumberFormatXml nf = new ExcelNumberFormatXml(_nameSpaceManager, n);
                    NumberFormats.Add(nf.Id, nf);
                    if (nf.NumFmtId >= NumberFormats.NextId) NumberFormats.NextId=nf.NumFmtId+1;
                }
            }

            //Fonts
            XmlNode fontNode = _styleXml.SelectSingleNode(FontsPath, _nameSpaceManager);
            foreach (XmlNode n in fontNode)
            {
                ExcelFontXml f = new ExcelFontXml(_nameSpaceManager, n);
                Fonts.Add(f.Id, f);
            }

            //Fills
            XmlNode fillNode = _styleXml.SelectSingleNode(FillsPath, _nameSpaceManager);
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
            XmlNode borderNode = _styleXml.SelectSingleNode(BordersPath, _nameSpaceManager);
            foreach (XmlNode n in borderNode)
            {
                ExcelBorderXml b = new ExcelBorderXml(_nameSpaceManager, n);
                Borders.Add(b.Id, b);
            }

            //cellStyleXfs
            XmlNode styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
            if (styleXfsNode != null)
            {
                foreach (XmlNode n in styleXfsNode)
                {
                    ExcelXfs item = new ExcelXfs(_nameSpaceManager, n, this);
                    CellStyleXfs.Add(item.Id, item);
                }
            }

            XmlNode styleNode = _styleXml.SelectSingleNode(CellXfsPath, _nameSpaceManager);
            for (int i = 0; i < styleNode.ChildNodes.Count; i++)
            {
                XmlNode n = styleNode.ChildNodes[i];
                ExcelXfs item = new ExcelXfs(_nameSpaceManager, n, this);
                CellXfs.Add(item.Id, item);
            }

            //cellStyle
            XmlNode namedStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _nameSpaceManager);
            if (namedStyleNode != null)
            {
                foreach (XmlNode n in namedStyleNode)
                {
                    ExcelNamedStyleXml item = new ExcelNamedStyleXml(_nameSpaceManager, n, this);
                    NamedStyles.Add(item.Name, item);
                }
            }

            //dxfsPath
            XmlNode dxfsNode = _styleXml.SelectSingleNode(dxfsPath, _nameSpaceManager);
            if (dxfsNode != null)
            {
                foreach (XmlNode x in dxfsNode)
                {
                    ExcelDxfStyleConditionalFormatting item = new ExcelDxfStyleConditionalFormatting(_nameSpaceManager, x, this);
                    Dxfs.Add(item.Id, item);
                }
            }
        }

        internal ExcelNamedStyleXml GetNormalStyle()
        {
            foreach (var style in NamedStyles)
            {
                if (style.BuildInId == 0)
                {
                    return style;
                }
            }
            if (_wb.Styles.NamedStyles.Count > 0)
            {
                return _wb.Styles.NamedStyles[0];
            }
            else
            {
                return null;
            }
        }

        internal ExcelStyle GetStyleObject(int Id,int PositionID, string Address)
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
            var rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
            var colCache = new Dictionary<int, ExcelValue>(address.End.Column - address.Start.Column + 1);
            var cellEnum = new CellStoreEnumerator<ExcelValue>(ws._values, address.Start.Row, address.Start.Column, address.End.Row, address.End.Column);
            var hasEnumValue=cellEnum.Next();
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
                                if (v._value==null)
                                {
                                    if (GetFromCache(colCache, col, ref s) == false)
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
                                    else
                                    {
                                        colCache.Add(col, new ExcelValue() { _styleId = 0 });
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
                        ws._values.SetValue(row,col,new ExcelValue { _value = value._value, _styleId = styleCashe[s] });
                    }
                    else
                    {
                        ExcelXfs st = CellXfs[s];
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
            while(!colCache.ContainsKey(--c))
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
                    while (cse.Next())
                    {
                        s = cse.Value._styleId;
                        if (s == 0) continue;
                        var c = ws.GetValueInner(cse.Row, cse.Column) as ExcelColumn;
                        if (c != null && c.ColumnMax < ExcelPackage.MaxColumns)
                        {
                            for (int col = c.ColumnMin; col < c.ColumnMax; col++)
                            {
                                if (!ws.ExistsStyleInner(rowNum, col))
                                {
                                    ws.SetStyleInner(rowNum, col, s);
                                }
                            }
                        }
                    }
                    ws.SetStyleInner(rowNum, 0, s);
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
            int v=0;
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
                        int r=0,c=col;
                        if(ws._values.PrevCell(ref r,ref c))
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

            int index = NamedStyles.FindIndexByID(e.Address);
            if (index >= 0)
            {
                int newId = CellStyleXfs[NamedStyles[index].StyleXfId].GetNewID(CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                int prevIx=NamedStyles[index].StyleXfId;
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
        /// Contain all named styles for that package
        /// </summary>
        public ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = new ExcelStyleCollection<ExcelNamedStyleXml>();
        /// <summary>
        /// Contain all differential formatting styles for the package
        /// </summary>
        public ExcelStyleCollection<ExcelDxfStyleConditionalFormatting> Dxfs = new ExcelStyleCollection<ExcelDxfStyleConditionalFormatting>();
        
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
            if (Template == null)
            {
                xfIdCopy = 0;
                positionID = -1;
                styles = this;
            }
            else
            {
                if (Template.PositionID < 0 && Template.Styles==this)
                {
                    xfIdCopy = Template.Index;
                    
                    positionID=Template.PositionID;
                    styles = this;
                }
                else
                {
                    xfIdCopy = Template.XfId;
                    positionID = -1;
                    styles = Template.Styles;
                }
            }
            //Clone namedstyle
            int styleXfId = CloneStyle(styles, xfIdCopy, true);
            //Close cells style
            CellStyleXfs[styleXfId].XfId = CellStyleXfs.Count-1;
            int xfid = CloneStyle(styles, xfIdCopy, true, true); //Always add a new style (We create a new named style here)
            CellXfs[xfid].XfId = styleXfId;
            style.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, styleXfId);
            style.StyleXfId = styleXfId;
            
            style.Name = name;
            int ix =_wb.Styles.NamedStyles.Add(style.Name, style);
            style.Style.SetIndex(ix);
            return style;
        }
        /// <summary>
        /// Update the changes to the Style.Xml file inside the package.
        /// This will remove any unused styles from the collections.
        /// </summary>
        public void UpdateXml()
        {
            RemoveUnusedStyles();

            //NumberFormat
            XmlNode nfNode=_styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
            if (nfNode == null)
            {
                CreateNode(NumberFormatsPath, true);
                nfNode = _styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
            }
            else
            {
                nfNode.RemoveAll();                
            }

            int count = 0;
            int normalIx = NamedStyles.FindIndexByBuildInId(0);
            if(normalIx<0)
            {
                normalIx = NamedStyles.FindIndexByID("normal");
            }
            if (NamedStyles.Count > 0 && normalIx>=0 && NamedStyles[normalIx].Style.Numberformat.NumFmtID >= 164)
            {
                ExcelNumberFormatXml nf = NumberFormats[NumberFormats.FindIndexByID(NamedStyles[normalIx].Style.Numberformat.Id)];
                nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                nf.newID = count++;
            }
            foreach (ExcelNumberFormatXml nf in NumberFormats)
            {
                if(!nf.BuildIn /*&& nf.newID<0*/) //Buildin formats are not updated.
                {
                    nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                    nf.newID = count;
                    count++;
                }
            }
            (nfNode as XmlElement).SetAttribute("count", count.ToString());

            //Font
            count=0;
            XmlNode fntNode = _styleXml.SelectSingleNode(FontsPath, _nameSpaceManager);
            fntNode.RemoveAll();

            //Normal should be first in the collection
            if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Font.Index > 0)
            {
                ExcelFontXml fnt = Fonts[NamedStyles[normalIx].Style.Font.Index];
                fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                fnt.newID = count++;
            }

            foreach (ExcelFontXml fnt in Fonts)
            {
                if (fnt.useCnt > 0/* && fnt.newID<0*/)
                {
                    fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                    fnt.newID = count;
                    count++;
                }
            }
            (fntNode as XmlElement).SetAttribute("count", count.ToString());


            //Fills
            count = 0;
            XmlNode fillsNode = _styleXml.SelectSingleNode(FillsPath, _nameSpaceManager);
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

            //Borders
            count = 0;
            XmlNode bordersNode = _styleXml.SelectSingleNode(BordersPath, _nameSpaceManager);
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

            XmlNode styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
            if (styleXfsNode == null && NamedStyles.Count > 0)
            {
                CreateNode(CellStyleXfsPath);
                styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
            }
            if (NamedStyles.Count > 0)
            {
                styleXfsNode.RemoveAll();
            }
            //NamedStyles
            count = normalIx > -1 ? 1 : 0;  //If we have a normal style, we make sure it's added first.

            XmlNode cellStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _nameSpaceManager);
            if(cellStyleNode!=null)
            {
                cellStyleNode.RemoveAll();
            }
            XmlNode cellXfsNode = _styleXml.SelectSingleNode(CellXfsPath, _nameSpaceManager);
            cellXfsNode.RemoveAll();

            if (NamedStyles.Count > 0 && normalIx >= 0)
            {
                NamedStyles[normalIx].newID = 0;
                AddNamedStyle(0, styleXfsNode, cellXfsNode, NamedStyles[normalIx]);
            }
            foreach (ExcelNamedStyleXml style in NamedStyles)
            {
                if (!style.Name.Equals("normal", StringComparison.OrdinalIgnoreCase))
                {
                    AddNamedStyle(count++, styleXfsNode, cellXfsNode, style);
                }
                else
                {
                    style.newID = 0;
                }
                cellStyleNode.AppendChild(style.CreateXmlNode(_styleXml.CreateElement("cellStyle", ExcelPackage.schemaMain)));
            }
            if (cellStyleNode!=null) (cellStyleNode as XmlElement).SetAttribute("count", count.ToString());
            if (styleXfsNode != null) (styleXfsNode as XmlElement).SetAttribute("count", count.ToString());

            //CellStyle
            int xfix = 0;
            foreach (ExcelXfs xf in CellXfs)
            {
                if (xf.useCnt > 0 && !(normalIx >= 0 && NamedStyles[normalIx].StyleXfId == xfix))
                {
                    cellXfsNode.AppendChild(xf.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                    xf.newID = count;
                    count++;
                }
                xfix++;
            }
            (cellXfsNode as XmlElement).SetAttribute("count", count.ToString());

            //Set dxf styling for conditional Formatting
            XmlNode dxfsNode = _styleXml.SelectSingleNode(dxfsPath, _nameSpaceManager);
            foreach (var ws in _wb.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                foreach (var cf in ws.ConditionalFormatting)
                {
                    if (cf.Style.HasValue)
                    {
                        int ix = Dxfs.FindIndexByID(cf.Style.Id);
                        if (ix < 0)
                        {
                            ((ExcelConditionalFormattingRule)cf).DxfId = Dxfs.Count;
                            Dxfs.Add(cf.Style.Id, cf.Style);
                            var elem = ((XmlDocument)TopNode).CreateElement("d", "dxf", ExcelPackage.schemaMain);
                            cf.Style.CreateNodes(new XmlHelperInstance(NameSpaceManager, elem), "");
                            dxfsNode.AppendChild(elem);
                        }
                        else
                        {
                            ((ExcelConditionalFormattingRule)cf).DxfId = ix;
                        }
                    }
                }
            }
            if (dxfsNode != null) (dxfsNode as XmlElement).SetAttribute("count", Dxfs.Count.ToString());
        }

        private void AddNamedStyle(int id, XmlNode styleXfsNode,XmlNode cellXfsNode, ExcelNamedStyleXml style)
        {
            var styleXfs = CellStyleXfs[style.StyleXfId];
            styleXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
            styleXfs.newID = id;
            styleXfs.XfId = style.StyleXfId;

            var ix = CellXfs.FindIndexByID(styleXfs.Id);
            if (ix < 0)
            {
                cellXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
            }
            else
            {
                if(id<0) CellXfs[ix].XfId = id;
                cellXfsNode.AppendChild(CellXfs[ix].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                CellXfs[ix].useCnt = 0;
                CellXfs[ix].newID = id;
            }

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
            int i = NamedStyles.FindIndexByID(Name);
            if (i >= 0)
            {
                int id = NamedStyles[i].XfId;
                if (id < 0)
                {
                    int styleXfId=NamedStyles[i].StyleXfId;
                    ExcelXfs newStyle = CellStyleXfs[styleXfId].Copy();
                    newStyle.XfId = styleXfId;
                    id = CellXfs.FindIndexByID(newStyle.Id);
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
            return CloneStyle(style, styleID, false, false);
        }
        internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle)
        {
            return CloneStyle(style, styleID, isNamedStyle, false);
        }
        internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle, bool allwaysAddCellXfs)
        {
            ExcelXfs xfs;
            lock (style)
            {
                if (isNamedStyle)
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
                        int ix = NumberFormats.FindIndexByID(format);
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
                    int ix = Fonts.FindIndexByID(xfs.Font.Id);
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
                    int ix = Borders.FindIndexByID(xfs.Border.Id);
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
                    int ix = Fills.FindIndexByID(xfs.Fill.Id);
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
                    var id = style.CellStyleXfs[xfs.XfId].Id;
                    var newId = CellStyleXfs.FindIndexByID(id);
                    if (newId >= 0)
                    {
                        newXfs.XfId = newId;
                    }
                    else if(style._wb!=_wb && allwaysAddCellXfs==false) //Not the same workbook, copy the namedstyle to the workbook or match the id
                    {
                        var nsFind = style.NamedStyles.ToDictionary(d => (d.StyleXfId));
                        if (nsFind.ContainsKey(xfs.XfId))
                        {
                            var st = nsFind[xfs.XfId];
                            if (NamedStyles.ExistsKey(st.Name))
                            {
                                newXfs.XfId = NamedStyles.FindIndexByID(st.Name);
                            }
                            else
                            {
                                var ns = CreateNamedStyle(st.Name, st.Style);
                                newXfs.XfId = NamedStyles.Count - 1;
                            }
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
                        index = CellXfs.FindIndexByID(newXfs.Id);
                        if (index < 0)
                        {
                            index = CellXfs.Add(newXfs.Id, newXfs);
                        }
                    }
                }
                return index;
            }
        }
        internal int CloneDxfStyle(ExcelStyles style, int styleID)
        {
            var copy = style.Dxfs[styleID];
            var ix = Dxfs.FindIndexByID(copy.Id);
            if(ix<0)
            {
                var parent = GetNode(dxfsPath);
                var node = _styleXml.CreateElement("d:dxf", ExcelPackage.schemaMain);
                parent.AppendChild(node);
                node.InnerXml = copy._helper.TopNode.InnerXml;
                var dxf = new ExcelDxfStyleConditionalFormatting(_nameSpaceManager, node, this);
                Dxfs.Add(copy.Id, dxf);
                return Dxfs.Count - 1;
            }
            else
            {
                return ix;
            }
        }
    }
}
