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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.ExcelXMLWriter;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    [Flags]
    internal enum CellFlags
    {
        //Merged = 0x1,
        RichText = 0x2,
        SharedFormula = 0x4,
        ArrayFormula = 0x8,
        DataTableFormula = 0x10
    }
    /// <summary>
	/// Represents an Excel worksheet and provides access to its properties and methods
	/// </summary>
    public class ExcelWorksheet : XmlHelper, IEqualityComparer<ExcelWorksheet>, IDisposable, IPictureRelationDocument
    {
        internal enum FormulaType
        {
            Normal,
            Shared,
            Array,
            DataTable
        }
        internal class Formulas
        {
            public Formulas(ISourceCodeTokenizer tokenizer)
            {
                _tokenizer = tokenizer;
            }

            private ISourceCodeTokenizer _tokenizer;
            internal int Index { get; set; }
            internal string Address { get; set; }
            internal FormulaType FormulaType { get; set; }
            string _formula = "";
            public string Formula
            {
                get
                {
                    return _formula;
                }
                set
                {
                    if (_formula != value)
                    {
                        _formula = value;
                        Tokens = null;
                    }
                }
            }
            public int StartRow { get; set; }
            public int StartCol { get; set; }
            public bool FirstCellDeleted { get; set; }  //del1
            public bool SecondCellDeleted { get; set; } //del2

            public bool DataTableIsTwoDimesional { get; set; } //dt2D
            public bool IsDataTableRow { get; set; } //dtr
            public string R1CellAddress { get; set; } //r1
            public string R2CellAddress { get; set; } //r2
            internal IEnumerable<Token> Tokens { get; set; }

            internal void SetTokens(string worksheet)
            {
                if (Tokens == null)
                {
                    Tokens = _tokenizer.Tokenize(Formula, worksheet);
                }
            }
            internal string GetFormula(int row, int column, string worksheet)
            {
                if ((StartRow == row && StartCol == column))
                {
                    return Formula;
                }

                SetTokens(worksheet);
                string f = "";
                foreach (var token in Tokens)
                {
                    if (token.TokenTypeIsSet(TokenType.ExcelAddress))
                    {
                        var a = new ExcelFormulaAddress(token.Value, (ExcelWorksheet)null);
                        if (a.IsFullColumn)
                        {
                            if (a.IsFullRow)
                            {
                                f += token.Value;
                            }
                            else
                            {
                                f += a.GetOffset(0, column - StartCol, true);
                            }
                        }
                        else if (a.IsFullRow)
                        {
                            f += a.GetOffset(row - StartRow, 0, true);
                        }
                        else
                        {
                            if (a.Table != null)
                            {
                                f += token.Value;
                            }
                            else
                            {
                                f += a.GetOffset(row - StartRow, column - StartCol, true);
                            }
                        }
                    }
                    else
                    {
                        if (token.TokenTypeIsSet(TokenType.StringContent))
                        {
                            f += token.Value.Replace("\"", "\"\"");
                        }
                        else
                        {
                            f += token.Value;
                        }
                    }
                }
                return f;
            }
            internal Formulas Clone()
            {
                return new Formulas(_tokenizer)
                {
                    Index = Index,
                    Address = Address,
                    FormulaType = FormulaType,
                    Formula = Formula,
                    StartRow = StartRow,
                    StartCol = StartCol,
                    DataTableIsTwoDimesional = DataTableIsTwoDimesional,
                    IsDataTableRow = IsDataTableRow,
                    R1CellAddress = R1CellAddress,
                    R2CellAddress = R2CellAddress,
                    FirstCellDeleted = FirstCellDeleted,
                    SecondCellDeleted = SecondCellDeleted,
                };
            }
        }
        /// <summary>
        /// Keeps track of meta data referencing cells or values.
        /// </summary>
        internal struct MetaDataReference
        {
            internal int cm;
            internal int vm;
            internal bool aca;
            internal bool ca;
        }
        /// <summary>
        /// Removes all formulas within the entire worksheet, but keeps the calculated values.
        /// </summary>
        public void ClearFormulas()
        {
            if (Dimension == null) return;

            var formulaCells = new CellStoreEnumerator<object>(_formulas, Dimension.Start.Row, Dimension.Start.Column, Dimension.End.Row, Dimension.End.Column);
            while (formulaCells.Next())
            {
                formulaCells.Value = null;
            }
        }

        /// <summary>
        /// Removes all values of cells with formulas in the entire worksheet, but keeps the formulas.
        /// </summary>
        public void ClearFormulaValues()
        {
            var formulaCell = new CellStoreEnumerator<object>(_formulas, Dimension.Start.Row, Dimension.Start.Column, Dimension.End.Row, Dimension.End.Column);
            while (formulaCell.Next())
            {

                var val = _values.GetValue(formulaCell.Row, formulaCell.Column);
                val._value = null;
                _values.SetValue(formulaCell.Row, formulaCell.Column, val);
            }
        }

        /// <summary>
        /// Collection containing merged cell addresses
        /// </summary>
        public class MergeCellsCollection : IEnumerable<string>
        {
            internal MergeCellsCollection()
            {

            }
            internal CellStore<int> _cells = new CellStore<int>();
            internal List<string> _list = new List<string>();
            /// <summary>
            /// Indexer for the collection
            /// </summary>
            /// <param name="row">The Top row of the merged cells</param>
            /// <param name="column">The Left column of the merged cells</param>
            /// <returns></returns>
            public string this[int row, int column]
            {
                get
                {
                    int ix = -1;
                    if (_cells.Exists(row, column, ref ix) && ix >= 0 && ix < _list.Count)  //Fixes issue 15075
                    {
                        return _list[ix];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            /// <summary>
            /// Indexer for the collection
            /// </summary>
            /// <param name="index">The index in the collection</param>
            /// <returns></returns>
            public string this[int index]
            {
                get
                {
                    return _list[index];
                }
            }
            internal void Add(ExcelAddressBase address, bool doValidate)
            {
                int ix = 0;

                //Validate
                if (doValidate && Validate(address) == false)
                {
                    throw (new ArgumentException("Can't merge and already merged range"));
                }
                lock (this)
                {
                    ix = _list.Count;
                    _list.Add(address.Address);
                    SetIndex(address, ix);
                }
            }

            private bool Validate(ExcelAddressBase address)
            {
                int ix = 0;
                if (_cells.Exists(address._fromRow, address._fromCol, ref ix))
                {
                    if (ix >= 0 && ix < _list.Count && _list[ix] != null && address.Address == _list[ix])
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                var cse = new CellStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
                //cells
                while (cse.Next())
                {
                    return false;
                }
                //Entire column
                cse = new CellStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
                while (cse.Next())
                {
                    return false;
                }
                //Entire row
                cse = new CellStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
                while (cse.Next())
                {
                    return false;
                }
                return true;
            }

            internal void SetIndex(ExcelAddressBase address, int ix)
            {
                if (address._fromRow == 1 && address._toRow == ExcelPackage.MaxRows) //Entire row
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        _cells.SetValue(0, col, ix);
                    }
                }
                else if (address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns) //Entire row
                {
                    for (int row = address._fromRow; row <= address._toRow; row++)
                    {
                        _cells.SetValue(row, 0, ix);
                    }
                }
                else
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        for (int row = address._fromRow; row <= address._toRow; row++)
                        {
                            _cells.SetValue(row, col, ix);
                        }
                    }
                }
            }
            /// <summary>
            /// Number of items in the collection
            /// </summary>
            public int Count
            {
                get
                {
                    return _list.Count;
                }
            }
            #region IEnumerable<string> Members

            /// <summary>
            /// Gets the enumerator for the collection
            /// </summary>
            /// <returns>The enumerator</returns>
            public IEnumerator<string> GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion

            #region IEnumerable Members

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion
            internal void Clear(ExcelAddressBase Destination)
            {
                var cse = new CellStoreEnumerator<int>(_cells, Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
                var used = new HashSet<int>();
                while (cse.Next())
                {
                    var v = cse.Value;
                    if (!used.Contains(v) && _list[v] != null)
                    {
                        var adr = new ExcelAddressBase(_list[v]);
                        if (!(Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Inside || Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Equal))
                        {
                            throw (new InvalidOperationException(string.Format("Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {0}", adr._address)));
                        }
                        used.Add(v);
                    }
                }

                _cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
                foreach (var i in used)
                {
                    _list[i] = null;
                }
            }

            internal void CleanupMergedCells()
            {
                _list = _list.Where(x => x != null).ToList();
            }
        }
        internal CellStoreValue _values;
        internal CellStore<object> _formulas;
        internal FlagCellStore _flags;
        internal CellStore<List<Token>> _formulaTokens;

        internal CellStore<Uri> _hyperLinks;
        internal CellStore<int> _commentsStore;
        internal CellStore<int> _threadedCommentsStore;
        internal CellStore<int?> _dataValidationsStore;

        internal CellStore<MetaDataReference> _metadataStore;

        internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();
        internal RangeSorter _rangeSorter;
        internal int _minCol = ExcelPackage.MaxColumns;
        internal int _maxCol = 0;
        internal int _nextControlId;
        #region Worksheet Private Properties
        internal ExcelPackage _package;
        private Uri _worksheetUri;
        private string _name;
        private int _sheetID;
        private int _positionId;
        private string _relationshipID;
        private XmlDocument _worksheetXml;
        internal ExcelWorksheetView _sheetView;
        internal ExcelHeaderFooter _headerFooter;
        #endregion
        #region ExcelWorksheet Constructor
        /// <summary>
        /// A worksheet
        /// </summary>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="excelPackage">Package</param>
        /// <param name="relID">Relationship ID</param>
        /// <param name="uriWorksheet">URI</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="sheetID">Sheet id</param>
        /// <param name="positionID">Position</param>
        /// <param name="hide">hide</param>
        public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID,
                              Uri uriWorksheet, string sheetName, int sheetID, int positionID,
                              eWorkSheetHidden? hide) :
            base(ns, null)
        {
            SchemaNodeOrder = new string[] { "sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst" };
            _package = excelPackage;
            _relationshipID = relID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = sheetID;
            _positionId = positionID;

            if (hide.HasValue)
            {
                Hidden = hide.Value;
            }

            /**** Cellstore ****/
            _values = new CellStoreValue();
            _formulas = new CellStore<object>();
            _flags = new FlagCellStore();
            _metadataStore = new CellStore<MetaDataReference>();
            _commentsStore = new CellStore<int>();
            _threadedCommentsStore = new CellStore<int>();
            _formulaTokens = new CellStore<List<Token>>();
            _hyperLinks = new CellStore<Uri>();
            _dataValidationsStore = new CellStore<int?>();
            _nextControlId = (PositionId + 1) * 1024 + 1;
            _names = new ExcelNamedRangeCollection(Workbook, this);

            _rangeSorter = new RangeSorter(this);

            CreateXml();
            TopNode = _worksheetXml.DocumentElement;
            LoadComments();
            LoadThreadedComments();
        }
        internal void LoadComments()
        {
            CreateVmlCollection();
            _comments = new ExcelCommentCollection(_package, this, NameSpaceManager);
        }
        internal void LoadThreadedComments()
        {
            _threadedComments = new ExcelWorksheetThreadedComments(Workbook.ThreadedCommentPersons, this);
        }

        #endregion
        /// <summary>
        /// The Uri to the worksheet within the package
        /// </summary>
        internal Uri WorksheetUri { get { return (_worksheetUri); } }
        /// <summary>
        /// The Zip.ZipPackagePart for the worksheet within the package
        /// </summary>
        internal Packaging.ZipPackagePart Part { get { return (_package.ZipPackage.GetPart(WorksheetUri)); } }
        /// <summary>
        /// The ID for the worksheet's relationship with the workbook in the package
        /// </summary>
        internal string RelationshipId { get { return (_relationshipID); } }
        /// <summary>
        /// The unique identifier for the worksheet.
        /// </summary>
        internal int SheetId { get { return (_sheetID); } set { _sheetID = value; } }
        internal bool IsChartSheet { get; set; } = false;
        internal static bool NameNeedsApostrophes(string ws)
        {
            if (ws[0] >= '0' && ws[0] <= '9')
            {
                return true;
            }
            if (StartsWithR1C1(ws))
            {
                return true;
            }
            foreach (var c in ws)
            {
                if (!(char.IsLetterOrDigit(c) || c == '_'))
                    return true;
            }
            return false;
        }

        private static bool StartsWithR1C1(string ws)
        {
            if (ws[0] == 'c' || ws[0] == 'C' || ws[0] == 'r' || ws[0] == 'R')
            {
                int ix = 1;
                if (ws.StartsWith("rc", StringComparison.OrdinalIgnoreCase)) ix = 2;
                if (ws.Length > ix && (ws[ix] >= '0' && ws[ix] <= '9'))
                {
                    if (ws[ix] == '0')
                    {
                        for (int i = ix + 1; i < ws.Length; i++)
                        {
                            if (ws[i] != '0')
                            {
                                if (ws[i] >= '1' && ws[i] <= '9')
                                {
                                    return true;
                                }
                            }
                        }
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// The position of the worksheet.
        /// </summary>
        internal int PositionId { get { return (_positionId); } set { _positionId = value; } }
        internal int IndexInList
        {
            get
            {
                if (_package == null) return -1;
                return (_positionId - _package._worksheetAdd);
            }
        }
        #region Worksheet Public Properties
        /// <summary>
        /// The index in the worksheets collection
        /// </summary>
        public int Index { get { return (_positionId); } }
        const string AutoFilterPath = "d:autoFilter";
        /// <summary>
        /// Address for autofilter
        /// <seealso cref="ExcelRangeBase.AutoFilter" />        
        /// </summary>
        /// 
        const string SortStatePath = "d:sortState";
        /// <summary>
        /// The auto filter address. 
        /// null means no auto filter.
        /// </summary>
        public ExcelAddressBase AutoFilterAddress
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                string address = GetXmlNodeString($"{AutoFilterPath}/@ref");
                if (address == "")
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(address);
                }
            }
            internal set
            {
                CheckSheetTypeAndNotDisposed();
                if (value == null)
                {
                    DeleteAllNode($"{AutoFilterPath}/@ref");
                }
                else
                {
                    SetXmlNodeString($"{AutoFilterPath}/@ref", value.Address);
                }
            }
        }
        ExcelAutoFilter _autoFilter = null;
        /// <summary>
        /// Autofilter settings
        /// </summary>
        public ExcelAutoFilter AutoFilter
        {
            get
            {
                if (_autoFilter == null)
                {
                    CheckSheetTypeAndNotDisposed();
                    var node = _worksheetXml.SelectSingleNode($"//{AutoFilterPath}", NameSpaceManager);
                    if (node == null) return null;
                    _autoFilter = new ExcelAutoFilter(NameSpaceManager, node, this);
                }
                return _autoFilter;
            }
        }

        SortState _sortState = null;

        /// <summary>
        /// Sets the sort state
        /// </summary>
        public SortState SortState
        {
            get
            {
                if (_sortState == null)
                {
                    CheckSheetTypeAndNotDisposed();
                    var node = _worksheetXml.SelectSingleNode($"//{SortStatePath}", NameSpaceManager);
                    if (node == null) return null;
                    _sortState = new SortState(NameSpaceManager, node);
                }
                return _sortState;
            }
        }

        internal void CheckSheetTypeAndNotDisposed()
        {
            if (this is ExcelChartsheet)
            {
                throw (new NotSupportedException("This property or method is not supported for a Chartsheet"));
            }
            if (_positionId == -1 && _values == null)
            {
                throw new ObjectDisposedException("ExcelWorksheet", "Worksheet has been disposed");
            }
        }

        /// <summary>
        /// Returns a ExcelWorksheetView object that allows you to set the view state properties of the worksheet
        /// </summary>
        public ExcelWorksheetView View
        {
            get
            {
                if (_sheetView == null)
                {
                    XmlNode node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("d:sheetViews/d:sheetView");     //this one should always exist. but check anyway
                        node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
                    }
                    _sheetView = new ExcelWorksheetView(NameSpaceManager, node, this);
                }
                return (_sheetView);
            }
        }

        /// <summary>
        /// The worksheet's display name as it appears on the tab
        /// </summary>
        public string Name
        {
            get { return (_name); }
            set
            {
                if (value == _name) return;
                value = _package.Workbook.Worksheets.ValidateFixSheetName(value);
                foreach (var ws in Workbook.Worksheets)
                {
                    if (ws.PositionId != PositionId && ws.Name.Equals(value, StringComparison.OrdinalIgnoreCase))
                    {
                        throw (new ArgumentException("Worksheet name must be unique"));
                    }
                }
                _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@name", _sheetID), value);
                ChangeNames(value);

                _name = value;
            }
        }

        internal int GetColumnWidthPixels(int col, decimal mdw)
        {
            return ExcelColumn.ColumnWidthToPixels(GetColumnWidth(col + 1), mdw);
        }
        internal decimal GetColumnWidth(int col)
        {
            var column = GetColumn(col);
            if (column == null)   //Check that the column exists
            {
                return (decimal)DefaultColWidth;
            }
            else
            {
                return (decimal)Columns[col].Width;
            }
        }

        private void ChangeNames(string value)
        {
            //Renames name in this Worksheet;
            foreach (var n in Workbook.Names)
            {
                if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
                {
                    n.ChangeWorksheet(_name, value);
                }
            }
            foreach (var ws in Workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    foreach (var n in ws.Names)
                    {
                        if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
                        {
                            n.ChangeWorksheet(_name, value);
                        }
                    }
                    ws.UpdateSheetNameInFormulas(_name, value);
                }
            }
        }
        internal ExcelNamedRangeCollection _names;
        /// <summary>
        /// Provides access to named ranges
        /// </summary>
        public ExcelNamedRangeCollection Names
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return _names;
            }
        }
        /// <summary>
        /// Indicates if the worksheet is hidden in the workbook
        /// </summary>
        public eWorkSheetHidden Hidden
        {
            get
            {
                string state = _package.Workbook.GetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
                if (state == "hidden")
                {
                    return eWorkSheetHidden.Hidden;
                }
                else if (state == "veryHidden")
                {
                    return eWorkSheetHidden.VeryHidden;
                }
                return eWorkSheetHidden.Visible;
            }
            set
            {

                if (value == eWorkSheetHidden.Visible)
                {
                    _package.Workbook.DeleteNode(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
                }
                else
                {
                    string v;
                    v = value.ToString();
                    v = v.Substring(0, 1).ToLowerInvariant() + v.Substring(1);
                    _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID), v);
                    DeactivateTab();
                }
            }
        }

        internal double GetRowHeight(int row)
        {
            object o = null;
            if (ExistsValueInner(row, 0, ref o) && o != null)   //Check that the row exists
            {
                var internalRow = (RowInternal)o;
                if (internalRow.Hidden)
                {
                    return 0;
                }
                else if (internalRow.Height >= 0)
                {
                    return internalRow.Height;
                }
                else
                {
                    return GetRowHeightFromCellFonts(row);
                }
            }
            else
            {
                //The row exists, check largest font in row

                /**** Default row height is assumed here. Excel calcualtes the row height from the larges font on the line. The formula to this calculation is undocumented, so currently its implemented with constants... ****/
                return GetRowHeightFromCellFonts(row);
            }
        }
        internal double GetRowHeightPixels(int row)
        {
            return GetRowHeight(row) / 0.75;
        }
        Dictionary<int, double> _textHeights = new Dictionary<int, double>();
        private double GetRowHeightFromCellFonts(int row)
        {
            var dh = DefaultRowHeight;
            if (double.IsNaN(dh) || CustomHeight == false)
            {
                var height = dh;

                var cse = new CellStoreEnumerator<ExcelValue>(_values, row, 0, row, ExcelPackage.MaxColumns);
                var styles = Workbook.Styles;
                while (cse.Next())
                {
                    var xfs = styles.CellXfs[cse.Value._styleId];
                    var f = styles.Fonts[xfs.FontId];
                    double rh;
                    if (_textHeights.ContainsKey(cse.Value._styleId))
                    {
                        rh = _textHeights[cse.Value._styleId];
                    }
                    else
                    {
                        rh = ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
                        _textHeights.Add(cse.Value._styleId, rh);
                    }

                    if (rh > height)
                    {
                        height = rh;
                    }
                }
                return height;
            }
            else
            {
                return dh;
            }
        }


        private void DeactivateTab()
        {
            if (PositionId == Workbook.View.ActiveTab)
            {
                var worksheets = Workbook.Worksheets;
                for (int i = PositionId + 1; i < worksheets.Count; i++)
                {
                    if (worksheets[i + _package._worksheetAdd].Hidden == eWorkSheetHidden.Visible)
                    {
                        Workbook.View.ActiveTab = i;
                        return;
                    }
                }
                for (int i = PositionId - 1; i >= 0; i--)
                {
                    if (worksheets[i + _package._worksheetAdd].Hidden == eWorkSheetHidden.Visible)
                    {
                        Workbook.View.ActiveTab = i;
                        return;
                    }
                }

            }
        }

        internal double _defaultRowHeight = double.NaN;
        /// <summary>
		/// Get/set the default height of all rows in the worksheet
		/// </summary>
        public double DefaultRowHeight
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (double.IsNaN(_defaultRowHeight))
                {
                    if (CustomHeight)
                    {
                        _defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
                    }
                    if (double.IsNaN(_defaultRowHeight))
                    {
                        _defaultRowHeight = GetRowHeightFromNormalStyle();
                    }
                }
                return _defaultRowHeight;
            }
            set
            {
                CheckSheetTypeAndNotDisposed();
                _defaultRowHeight = value;
                if (double.IsNaN(value))
                {
                    DeleteNode("d:sheetFormatPr/@defaultRowHeight");
                }
                else
                {
                    SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
                    //Check if this is the default width for the normal style
                    double defHeight = GetRowHeightFromNormalStyle();
                    CustomHeight = true;
                }
            }
        }
        /// <summary>
		/// If true, empty rows are hidden by default.
        /// This reduces the size of the package and increases performance if most of the rows in a worksheet are hidden.
		/// </summary>
        public bool RowZeroHeight
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return GetXmlNodeBool("d:sheetFormatPr/@zeroHeight");
            }
            set
            {
                CheckSheetTypeAndNotDisposed();

                SetXmlNodeBool("d:sheetFormatPr/@zeroHeight", value, false);
            }
        }

        private double GetRowHeightFromNormalStyle()
        {
            var ix = Workbook.Styles.GetNormalStyleIndex();
            if (ix >= 0)
            {
                var f = Workbook.Styles.NamedStyles[ix].Style.Font;
                if (f.Name.Equals("Calibri", StringComparison.OrdinalIgnoreCase) && f.Size == 11) //Default normal font
                {
                    return 15;
                }
                return ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
            }
            else
            {
                return 15;   //Default Calibri 11
            }
        }

        /// <summary>
        /// 'True' if defaultRowHeight value has been manually set, or is different from the default value.
        /// Is automaticlly set to 'True' when assigning the DefaultRowHeight property
        /// </summary>
        public bool CustomHeight
        {
            get
            {
                return GetXmlNodeBool("d:sheetFormatPr/@customHeight");
            }
            set
            {
                SetXmlNodeBool("d:sheetFormatPr/@customHeight", value);
            }
        }
        /// <summary>
        /// Get/set the default width of all columns in the worksheet
        /// </summary>
        public double DefaultColWidth
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                double ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
                if (double.IsNaN(ret))
                {
                    var mfw = Convert.ToDouble(Workbook.MaxFontWidth);
                    var margin = 5d;
                    var width = Math.Truncate((8 * mfw + margin) / mfw * 256d) / 256d;
                    var widthPx = Math.Truncate(((256d * width + Math.Truncate(128d / mfw)) / 256d) * mfw);
                    var widthPxAdj = widthPx + (8 - (widthPx % 8));

                    var styles = _package.Workbook.Styles;
                    var sub = Math.Truncate(widthPxAdj / 120);
                    return Math.Truncate(widthPxAdj / (mfw - sub) * 256d) / 256d;
                }
                return ret;
            }
            set
            {
                CheckSheetTypeAndNotDisposed();
                SetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth", value);

                if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
                {
                    SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", GetRowHeightFromNormalStyle().ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        /** <outlinePr applyStyles="1" summaryBelow="0" summaryRight="0" /> **/
        const string outLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow";
        /// <summary>
        /// If true, summary rows are showen below the details, otherwise above.
        /// </summary>
        public bool OutLineSummaryBelow
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return GetXmlNodeBool(outLineSummaryBelowPath);
            }
            set
            {
                CheckSheetTypeAndNotDisposed();
                SetXmlNodeString(outLineSummaryBelowPath, value ? "1" : "0");
            }
        }
        const string outLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";
        /// <summary>
        /// If true, summary columns are to right of details otherwise to the left.
        /// </summary>
        public bool OutLineSummaryRight
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return GetXmlNodeBool(outLineSummaryRightPath);
            }
            set
            {
                CheckSheetTypeAndNotDisposed();
                SetXmlNodeString(outLineSummaryRightPath, value ? "1" : "0");
            }
        }
        const string outLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";
        /// <summary>
        /// Automatic styles
        /// </summary>
        public bool OutLineApplyStyle
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return GetXmlNodeBool(outLineApplyStylePath);
            }
            set
            {
                CheckSheetTypeAndNotDisposed();
                SetXmlNodeString(outLineApplyStylePath, value ? "1" : "0");
            }
        }
        const string tabColorPath = "d:sheetPr/d:tabColor/@rgb";
        /// <summary>
        /// Color of the sheet tab
        /// </summary>
        public Color TabColor
        {
            get
            {
                string col = GetXmlNodeString(tabColorPath);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                SetXmlNodeString(tabColorPath, value.ToArgb().ToString("X"));
            }
        }
        const string codeModuleNamePath = "d:sheetPr/@codeName";
        internal string CodeModuleName
        {
            get
            {
                return GetXmlNodeString(codeModuleNamePath);
            }
            set
            {
                SetXmlNodeString(codeModuleNamePath, value);
            }
        }
        internal void CodeNameChange(string value)
        {
            CodeModuleName = value;
        }
        /// <summary>
        /// The VBA code modul for the worksheet, if the package contains a VBA project.
        /// <seealso cref="ExcelWorkbook.CreateVBAProject"/>
        /// </summary>  
        public VBA.ExcelVBAModule CodeModule
        {
            get
            {
                if (_package.Workbook.VbaProject != null)
                {
                    return _package.Workbook.VbaProject.Modules[CodeModuleName];
                }
                else
                {
                    return null;
                }
            }
        }
        #region WorksheetXml
        /// <summary>
        /// The XML document holding the worksheet data.
        /// All column, row, cell, pagebreak, merged cell and hyperlink-data are loaded into memory and removed from the document when loading the document.        
        /// </summary>
        public XmlDocument WorksheetXml
        {
            get
            {
                return (_worksheetXml);
            }
        }
        internal ExcelVmlDrawingCollection _vmlDrawings = null;
        /// <summary>
        /// Vml drawings. underlaying object for comments
        /// </summary>
        internal ExcelVmlDrawingCollection VmlDrawings
        {
            get
            {
                if (_vmlDrawings == null)
                {
                    CreateVmlCollection();
                }
                return _vmlDrawings;
            }
        }
        internal ExcelCommentCollection _comments = null;
        /// <summary>
        /// Collection of comments
        /// </summary>
        public ExcelCommentCollection Comments
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return _comments;
            }
        }

        internal ExcelWorksheetThreadedComments _threadedComments = null;

        /// <summary>
        /// A collection of threaded comments referenced in the worksheet.
        /// </summary>
        public ExcelWorksheetThreadedComments ThreadedComments
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return _threadedComments;
            }
        }

        internal Uri ThreadedCommentsUri
        {
            get
            {
                var rel = Part.GetRelationshipsByType(ExcelPackage.schemaThreadedComment);
                if (rel != null && rel.Any())
                {
                    var uri = rel.First().TargetUri.OriginalString.Split('/').Last();
                    uri = "/xl/threadedComments/" + uri;
                    return new Uri(uri, UriKind.Relative);
                }
                return GetThreadedCommentUri();
            }
        }

        private Uri GetThreadedCommentUri()
        {
            var index = 1;
            var uri = new Uri("/xl/threadedComments/threadedComment" + index + ".xml", UriKind.Relative);
            uri = UriHelper.ResolvePartUri(Workbook.WorkbookUri, uri);
            while (Part.Package.PartExists(uri))
            {
                uri = new Uri("/xl/threadedComments/threadedComment" + (++index) + ".xml", UriKind.Relative);
                uri = UriHelper.ResolvePartUri(Workbook.WorkbookUri, uri);
            }

            return uri;
        }


        private void CreateVmlCollection()
        {
            var relIdNode = _worksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", NameSpaceManager);
            if (relIdNode == null)
            {
                _vmlDrawings = new ExcelVmlDrawingCollection(this, null);
            }
            else
            {
                if (Part.RelationshipExists(relIdNode.Value))
                {
                    var rel = Part.GetRelationship(relIdNode.Value);
                    var vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                    _vmlDrawings = new ExcelVmlDrawingCollection(this, vmlUri);
                    _vmlDrawings.RelId = rel.Id;
                }
            }
        }

        private void CreateXml()
        {
            _worksheetXml = new XmlDocument();
            _worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
            Packaging.ZipPackagePart packPart = _package.ZipPackage.GetPart(WorksheetUri);
            string xml = "";

            // First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...
            bool doAdjust = _package.DoAdjustDrawings;
            _package.DoAdjustDrawings = false;
            //bool isZipStream;
            WorksheetZipStream stream;
            if (packPart.Entry?.UncompressedSize > int.MaxValue)
            {
                MoveEntry(packPart.Package._zip, packPart.Entry);
                stream = new WorksheetZipStream(packPart.Package._zip, true, packPart.Entry.UncompressedSize);
                //isZipStream = true;
            }
            else
            {
                stream = new WorksheetZipStream(packPart.GetStream(), true);
                //isZipStream = false;
            }
#if Core
            var xr = XmlReader.Create(stream, new XmlReaderSettings()
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreWhitespace = true
            });
#else
            var xr = new XmlTextReader(stream);
#if NET35
            xr.ProhibitDtd = true;
#else
            xr.DtdProcessing = DtdProcessing.Prohibit;
#endif
            xr.WhitespaceHandling = WhitespaceHandling.None;
#endif            
            LoadColumns(xr);    //columnXml
            var lastXmlElement = "sheetData";
            xml = stream.GetBufferAsStringRemovingElement(false, lastXmlElement);
            long start = stream.Position;
            LoadCells(xr);
            var nextElementLength = GetAttributeLength(xr);
            stream.SetWriteToBuffer();
            
            LoadMergeCells(xr);
            var nextElement = "dataValidations";
            if (xr.ReadUntil(1, NodeOrders.WorksheetTopElementOrder, nextElement))
            {
                xml = stream.ReadFromEndElement(lastXmlElement, xml, nextElement, false, xr.Prefix);
                LoadDataValidations(xr);
                stream.SetWriteToBuffer();
                lastXmlElement = nextElement;
            }
            LoadHyperLinks(xr);
            LoadRowPageBreakes(xr);
            LoadColPageBreakes(xr);
            nextElement = "extLst";
            if (xr.ReadUntil(1, NodeOrders.WorksheetTopElementOrder, nextElement))
            {
                LoadExtLst(xr, stream, ref xml, ref lastXmlElement);
            }
            if(!string.IsNullOrEmpty(lastXmlElement))
            {
                xml = stream.ReadFromEndElement(lastXmlElement, xml);
            }


            // now release stream buffer (already converted whole Xml into XmlDocument Object and String)
            stream.Dispose();
            packPart.Stream = RecyclableMemory.GetStream();

            Encoding encoding = Encoding.UTF8;
            //first char is invalid sometimes?? 
            if (xml[0] != '<')
                LoadXmlSafe(_worksheetXml, xml.Substring(1, xml.Length - 1), encoding);
            else
                LoadXmlSafe(_worksheetXml, xml, encoding);

            _package.DoAdjustDrawings = doAdjust;
            ClearNodes();
        }

        private static long GetPosition(ZipPackagePart packPart, bool isZipStream, WorksheetZipStream stream, int nextElementLength)
        {
            if (isZipStream)
            {
                return packPart.Entry.ArchiveStream.Position;
            }
            else
            {
                return stream.Position - nextElementLength;
            }

        }

        private void MoveEntry(ZipInputStream zip, ZipEntry entry)
        {
            zip.Position = 0;
            ZipEntry e;
            do
            {
                e = zip.GetNextEntry();
            }
            while (e.FileDataPosition != entry.FileDataPosition);
        }

        /// <summary>
        /// Get the lenth of the attributes
        /// Conditional formatting attributes can be extremly long som get length of the attributes to finetune position.
        /// </summary>
        /// <param name="xr"></param>
        /// <returns></returns>
        private int GetAttributeLength(XmlReader xr)
        {
            if (xr.NodeType != XmlNodeType.Element) return 0;
            var length = 0;

            for (int i = 0; i < xr.AttributeCount; i++)
            {
                var a = xr.GetAttribute(i);
                length += string.IsNullOrEmpty(a) ? 0 : a.Length;
            }
            return length;
        }
        private void LoadRowPageBreakes(XmlReader xr)
        {
            if (!xr.ReadUntil(1, "rowBreaks", "colBreaks", "extLst")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "brk")
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int id;
                        if (int.TryParse(xr.GetAttribute("id"), NumberStyles.Number, CultureInfo.InvariantCulture, out id))
                        {
                            Row(id).PageBreak = true;
                        }
                    }
                }
                else
                {
                    break;
                }
            }
        }
        private void LoadColPageBreakes(XmlReader xr)
        {
            if (!xr.ReadUntil(1, "colBreaks", "extLst")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "brk")
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int id;
                        if (int.TryParse(xr.GetAttribute("id"), NumberStyles.Number, CultureInfo.InvariantCulture, out id))
                        {
                            Column(id).PageBreak = true;
                        }
                    }
                }
                else
                {
                    break;
                }
            }
        }

        private void ClearNodes()
        {
            if (_worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager).RemoveAll();
            }
        }
        const int BLOCKSIZE = 8192;


        private void LoadColumns(XmlReader xr)//(string xml)
        {
            if (xr.ReadUntil(1, "cols", "sheetData"))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Whitespace) continue;
                    if (xr.LocalName != "col") break;
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int min = int.Parse(xr.GetAttribute("min"));

                        ExcelColumn col = new ExcelColumn(this, min);

                        col.ColumnMax = int.Parse(xr.GetAttribute("max"));
                        col.Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture);
                        col.BestFit = GetBoolFromString(xr.GetAttribute("bestFit"));
                        col.Collapsed = GetBoolFromString(xr.GetAttribute("collapsed"));
                        col.Phonetic = GetBoolFromString(xr.GetAttribute("phonetic"));
                        col.OutlineLevel = (short)(xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture));
                        col.Hidden = GetBoolFromString(xr.GetAttribute("hidden"));
                        SetValueInner(0, min, col);

                        int style;
                        if (!(xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), NumberStyles.Number, CultureInfo.InvariantCulture, out style)))
                        {
                            SetStyleInner(0, min, style);
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Load Hyperlinks
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadHyperLinks(XmlReader xr)
        {
            if (!xr.ReadUntil(1, "hyperlinks", "rowBreaks", "colBreaks", "extLst")) return;
            var delRelIds = new HashSet<string>();
            while (xr.Read())
            {
                if (xr.LocalName == "hyperlink")
                {
                    if (xr.NodeType == XmlNodeType.EndElement) continue;
                    int fromRow, fromCol, toRow, toCol;
                    var reference = xr.GetAttribute("ref");
                    if (reference != null && ExcelCellBase.IsValidAddress(reference))
                    {
                        ExcelCellBase.GetRowColFromAddress(xr.GetAttribute("ref"), out fromRow, out fromCol, out toRow, out toCol);
                        ExcelHyperLink hl = null;
                        if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
                        {
                            var rId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
                            var rel = Part.GetRelationship(rId);

                            if (rel.TargetUri == null)
                            {
                                if (rel.Target.StartsWith("#", StringComparison.OrdinalIgnoreCase) && ExcelCellBase.IsValidAddress(rel.Target.Substring(1)))
                                {
                                    var a = new ExcelAddressBase(rel.Target.Substring(1));
                                    hl = new ExcelHyperLink(a.FullAddress, string.IsNullOrEmpty(a.WorkSheetName) ? a.Address : a.WorkSheetName);
                                }
                            }
                            else
                            {
                                var uri = rel.TargetUri;
                                if (uri.IsAbsoluteUri)
                                {
                                    try
                                    {
                                        hl = new ExcelHyperLink(uri.AbsoluteUri);
                                    }
                                    catch
                                    {
                                        hl = new ExcelHyperLink(uri.OriginalString, UriKind.Absolute);
                                    }
                                }
                                else
                                {
                                    hl = new ExcelHyperLink(uri.OriginalString, UriKind.Relative);
                                }
                            }
                            hl.Target = rel.Target;
                            hl.RId = rId;
                        }
                        else if (xr.GetAttribute("location") != null)
                        {
                            hl = GetHyperlinkFromRef(xr, "location", fromRow, toRow, fromCol, toCol);
                        }
                        else if (xr.GetAttribute("ref") != null)
                        {
                            hl = GetHyperlinkFromRef(xr, "ref", fromRow, toRow, fromCol, toCol);
                        }
                        else
                        {
                            // not enough info to create a hyperlink, move to next.
                            continue;
                        }

                        string tt = xr.GetAttribute("tooltip");
                        if (!string.IsNullOrEmpty(tt))
                        {
                            hl.ToolTip = tt;
                        }
                        for (int row = fromRow; row <= toRow; row++)
                        {
                            for (int col = fromCol; col <= toCol; col++)
                            {
                                _hyperLinks.SetValue(row, col, hl);
                            }
                        }
                        if (string.IsNullOrEmpty(hl.RId) == false && delRelIds.Contains(hl.RId) == false) delRelIds.Add(hl.RId);
                    }
                }
                else
                {
                    if (xr.NodeType == XmlNodeType.Whitespace) continue;
                    break;
                }
            }
            delRelIds.ToList().ForEach(x => Part.DeleteRelationship(x));
        }
        internal ExcelRichTextCollection GetRichText(int row, int col, ExcelRangeBase r)
        {
            XmlDocument xml = new XmlDocument();
            var v = GetValueInner(row, col);
            var isRt = _flags.GetFlagValue(row, col, CellFlags.RichText);
            if (v != null)
            {
                if (isRt)
                {
                    XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + v.ToString() + "</d:si>", Encoding.UTF8);
                }
                else
                {
                    xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + OfficeOpenXml.Utils.ConvertUtil.ExcelEscapeString(v.ToString()) + "</d:t></d:r></d:si>");
                }
            }
            else
            {
                xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
            }
            if (r == null)
            {
                return new ExcelRichTextCollection(NameSpaceManager, xml.SelectSingleNode("d:si", NameSpaceManager), this);
            }
            else
            {
                return new ExcelRichTextCollection(NameSpaceManager, xml.SelectSingleNode("d:si", NameSpaceManager), r);
            }
        }

        private ExcelHyperLink GetHyperlinkFromRef(XmlReader xr, string refTag, int fromRow = 0, int toRow = 0, int fromCol = 0, int toCol = 0)
        {
            var hl = new ExcelHyperLink(xr.GetAttribute(refTag), xr.GetAttribute("display"));
            hl.RowSpann = toRow - fromRow;
            hl.ColSpann = toCol - fromCol;
            return hl;
        }
        internal ExcelDataValidationCollection _dataValidations = null;
        /// <summary>
        /// DataValidation defined in the worksheet. Use the Add methods to create DataValidations and add them to the worksheet. Then
        /// set the properties on the instance returned.
        /// Must know worksheet or at least worksheet name to determine if extLst when user input DataValidations in API.
        /// </summary>
        /// <seealso cref="ExcelDataValidationCollection"/>
        public ExcelDataValidationCollection DataValidations
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (_dataValidations == null)
                {
                    _dataValidations = new ExcelDataValidationCollection(this);
                }
                return _dataValidations;
            }
        }

        private void LoadDataValidations(XmlReader xr)
        {
            _dataValidations = new ExcelDataValidationCollection(xr, this);
        }

        private void LoadExtLst(XmlReader xr, WorksheetZipStream stream, ref string xml, ref string lastXmlElement)
        {
            string lastUri = "";
            while (xr.ReadUntil(2, "ext"))
            {
                if (xr.GetAttribute("uri") == ExtLstUris.DataValidationsUri)
                {
                    xml = stream.ReadToExt(xml, ExtLstUris.DataValidationsUri, ref lastXmlElement, lastUri);
                    lastUri = ExtLstUris.DataValidationsUri;
                    stream.WriteToBuffer = false;
                    xr.Read();

                    if (_dataValidations == null)
                    {
                        _dataValidations = new ExcelDataValidationCollection(xr, this);
                    }
                    else
                    {
                        _dataValidations.ReadDataValidations(xr);
                    }
                    xr.Read(); //Read over ext end tag

                    stream.SetWriteToBuffer();
                }
                else
                {
                    //TODO: add other extLst options here. For now avoid infinite loop.
                    xr.Read();
                }
            }
            if(string.IsNullOrEmpty(lastUri)==false)
            {
                stream.ReadToEnd();
                xml = stream.ReadToExt(xml, "", ref lastXmlElement, lastUri);
            }
        }

        /// <summary>
        /// Load cells
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadCells(XmlReader xr)
        {
            xr.ReadUntil(1, "sheetData", "dataValidations", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks", "extLst");
            ExcelAddressBase address = null;
            string type = "";
            int style = 0;
            int row = 0;
            int col = 0;
            xr.Read();

            while (!xr.EOF)
            {
                while (xr.NodeType == XmlNodeType.EndElement || xr.NodeType == XmlNodeType.None)
                {
                    xr.Read();
                    if (xr.EOF) return;
                    continue;
                }
                if (xr.LocalName == "row")
                {
                    col = 0;
                    var r = xr.GetAttribute("r");
                    if (r == null)
                    {
                        row++;
                    }
                    else
                    {
                        row = Convert.ToInt32(r);
                    }

                    if (DoAddRow(xr))
                    {
                        SetValueInner(row, 0, AddRow(xr, row));
                        if (xr.GetAttribute("s") != null)
                        {
                            var styleId = int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture);
                            SetStyleInner(row, 0, (styleId < 0 ? 0 : styleId));
                        }
                    }
                    xr.Read();
                }
                else if (xr.LocalName == "c")
                {
                    var r = xr.GetAttribute("r");
                    if (r == null)
                    {
                        //Handle cells with no reference
                        col++;
                        address = new ExcelAddressBase(row, col, row, col);
                    }
                    else
                    {
                        address = new ExcelAddressBase(r);
                        col = address._fromCol;
                    }


                    //Datetype
                    if (xr.GetAttribute("t") != null)
                    {
                        type = xr.GetAttribute("t");
                        //_types.SetValue(address._fromRow, address._fromCol, type); 
                    }
                    else
                    {
                        type = "";
                    }
                    //Style
                    if (xr.GetAttribute("s") != null)
                    {
                        style = int.Parse(xr.GetAttribute("s"));
                        SetStyleInner(address._fromRow, address._fromCol, style < 0 ? 0 : style);
                        //SetValueInner(address._fromRow, address._fromCol, null); //TODO:Better Performance ??
                    }
                    else
                    {
                        style = 0;
                    }
                    //Meta data. Meta data is only preserved by EPPlus at this point
                    var cm = xr.GetAttribute("cm");
                    var vm = xr.GetAttribute("vm");
                    if (cm != null || vm != null)
                    {
                        _metadataStore.SetValue(
                            address._fromRow,
                            address._fromCol,
                            new MetaDataReference()
                            {
                                cm = string.IsNullOrEmpty(cm) ? 0 : int.Parse(cm),
                                vm = string.IsNullOrEmpty(vm) ? 0 : int.Parse(vm)
                            });
                    };

                    xr.Read();
                }
                else if (xr.LocalName == "v")
                {
                    SetValueFromXml(xr, type, style, address._fromRow, address._fromCol);

                    xr.Read();
                }
                else if (xr.LocalName == "f")
                {
                    string t = xr.GetAttribute("t");

                    var aca = xr.GetAttribute("aca");
                    var ca = xr.GetAttribute("ca");
                    //Meta data and formula settings. Meta data is only preserved by EPPlus at this point
                    if (aca != null || ca != null)
                    {
                        var md = _metadataStore.GetValue(row, col);
                        md.aca = aca == "1";
                        md.ca = ca == "1";
                        _metadataStore.SetValue(
                            row,
                            col,
                            md);
                    }

                    if (t == null || t == "normal")
                    {
                        var formula = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
                        if (!string.IsNullOrEmpty(formula))
                        {
                            _formulas.SetValue(address._fromRow, address._fromCol, formula);
                        }
                        SetValueInner(address._fromRow, address._fromCol, null);
                    }
                    else if (t == "shared")
                    {

                        string si = xr.GetAttribute("si");
                        if (si != null)
                        {
                            var sfIndex = int.Parse(si);
                            _formulas.SetValue(address._fromRow, address._fromCol, sfIndex);
                            SetValueInner(address._fromRow, address._fromCol, null);
                            string fAddress = xr.GetAttribute("ref");
                            string formula = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
                            if (formula != "")
                            {
                                _sharedFormulas.Add(sfIndex, new Formulas(SourceCodeTokenizer.Default) { Index = sfIndex, Formula = formula, Address = fAddress, StartRow = address._fromRow, StartCol = address._fromCol, FormulaType = FormulaType.Shared });
                            }
                        }
                        else
                        {
                            xr.Read();  //Something is wrong in the sheet, read next
                        }
                    }
                    else if (t == "array")
                    {
                        string refAddress = xr.GetAttribute("ref");
                        string formula = xr.ReadElementContentAsString();
                        var afIndex = GetMaxShareFunctionIndex(true);
                        if (!string.IsNullOrEmpty(refAddress))
                        {
                            WriteArrayFormulaRange(refAddress, afIndex, CellFlags.ArrayFormula);
                        }

                        _sharedFormulas.Add(afIndex, new Formulas(SourceCodeTokenizer.Default) { Index = afIndex, Formula = formula, Address = refAddress, StartRow = address._fromRow, StartCol = address._fromCol, FormulaType = FormulaType.Array });
                    }
                    else if (t == "dataTable")
                    {
                        var afIndex = GetMaxShareFunctionIndex(true);
                        string refAddress = xr.GetAttribute("ref");
                        var f = new Formulas(SourceCodeTokenizer.Default)
                        {
                            Index = afIndex,
                            Address = refAddress,
                            StartRow = address._fromRow,
                            StartCol = address._fromCol,
                            FormulaType = FormulaType.DataTable,
                            FirstCellDeleted = GetBoolFromString(xr.GetAttribute("del1")),
                            SecondCellDeleted = GetBoolFromString(xr.GetAttribute("del2")),
                            DataTableIsTwoDimesional = GetBoolFromString(xr.GetAttribute("dt2D")),
                            R1CellAddress = xr.GetAttribute("r1") ?? "",
                            R2CellAddress = xr.GetAttribute("r2") ?? ""
                        };
                        f.Formula = xr.ReadElementContentAsString();
                        if (!string.IsNullOrEmpty(refAddress))
                        {
                            WriteArrayFormulaRange(refAddress, afIndex, CellFlags.DataTableFormula);
                        }

                        _sharedFormulas.Add(afIndex, f);
                        //xr.Read();
                    }
                    else // ??? some other type
                    {
                        xr.Read();  //Something is wrong in the sheet, read next
                    }

                }
                else if (xr.LocalName == "is")   //Inline string
                {
                    xr.Read();
                    if (xr.LocalName == "t")
                    {
                        SetValueInner(address._fromRow, address._fromCol, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
                    }
                    else
                    {
                        if (xr.LocalName == "r")
                        {
                            var rXml = xr.ReadOuterXml();
                            while (xr.LocalName == "r")
                            {
                                rXml += xr.ReadOuterXml();
                            }
                            SetValueInner(address._fromRow, address._fromCol, rXml);
                        }
                        else
                        {
                            SetValueInner(address._fromRow, address._fromCol, xr.ReadOuterXml());
                        }
                        _flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.RichText);
                    }
                }
                else
                {
                    break;
                }
            }
        }
        private void WriteArrayFormulaRange(string address, int index, CellFlags type)
        {
            var refAddress = new ExcelAddressBase(address);
            for (int r = refAddress._fromRow; r <= refAddress._toRow; r++)
            {
                for (int c = refAddress._fromCol; c <= refAddress._toCol; c++)
                {
                    _formulas.SetValue(r, c, index);
                    SetValueInner(r, c, null);
                    _flags.SetFlagValue(r, c, true, type);
                }
            }
        }
        private bool DoAddRow(XmlReader xr)
        {
            var c = xr.GetAttribute("r") == null ? 0 : 1;
            if (xr.GetAttribute("spans") != null)
            {
                c++;
            }
            return xr.AttributeCount > c;
        }
        /// <summary>
        /// Load merged cells
        /// </summary>
        /// <param name="xr"></param>
        private void LoadMergeCells(XmlReader xr)
        {
            if (xr.ReadUntil(1, "mergeCells", "dataValidations", "hyperlinks", "rowBreaks", "colBreaks", "extLst") && !xr.EOF)
            {
                while (xr.Read())
                {
                    if (xr.LocalName != "mergeCell") break;
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        string address = xr.GetAttribute("ref");
                        _mergedCells.Add(new ExcelAddress(address), false);
                    }
                }
            }
        }

        /// <summary>
        /// Reads a row from the XML reader
        /// </summary>
        /// <param name="xr">The reader</param>
        /// <param name="row">The row number</param>
        /// <returns></returns>
        private RowInternal AddRow(XmlReader xr, int row)
        {
            return new RowInternal()
            {
                Collapsed = GetBoolFromString(xr.GetAttribute("collapsed")),
                OutlineLevel = (xr.GetAttribute("outlineLevel") == null ? (short)0 : short.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture)),
                Height = (xr.GetAttribute("ht") == null ? -1 : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture)),
                Hidden = GetBoolFromString(xr.GetAttribute("hidden")),
                Phonetic = GetBoolFromString(xr.GetAttribute("ph")),
                CustomHeight = GetBoolFromString(xr.GetAttribute("customHeight"))
            };
        }

        private void SetValueFromXml(XmlReader xr, string type, int styleID, int row, int col)
        {
            var v = ConvertUtil.GetValueFromType(xr, type, styleID, Workbook);
            if (type == "s" && v is int ix)
            {
                SetValueInner(row, col, _package.Workbook._sharedStringsList[ix].Text);
                if (_package.Workbook._sharedStringsList[ix].isRichText)
                {
                    _flags.SetFlagValue(row, col, true, CellFlags.RichText);
                }
            }
            else
            {
                SetValueInner(row, col, v);
            }
        }

        //private string GetSharedString(int stringID)
        //{
        //    string retValue = null;
        //    XmlNodeList stringNodes = xlPackage.Workbook.SharedStringsXml.SelectNodes(string.Format("//d:si", stringID), NameSpaceManager);
        //    XmlNode stringNode = stringNodes[stringID];
        //    if (stringNode != null)
        //        retValue = stringNode.InnerText;
        //    return (retValue);
        //}
        #endregion
        #region HeaderFooter
        /// <summary>
        /// A reference to the header and footer class which allows you to 
        /// set the header and footer for all odd, even and first pages of the worksheet
        /// </summary>
        /// <remarks>
        /// To format the text you can use the following format
        /// <list type="table">
        /// <listheader><term>Prefix</term><description>Description</description></listheader>
        /// <item><term>&amp;U</term><description>Underlined</description></item>
        /// <item><term>&amp;E</term><description>Double Underline</description></item>
        /// <item><term>&amp;K:xxxxxx</term><description>Color. ex &amp;K:FF0000 for red</description></item>
        /// <item><term>&amp;"Font,Regular Bold Italic"</term><description>Changes the font. Regular or Bold or Italic or Bold Italic can be used. ex &amp;"Arial,Bold Italic"</description></item>
        /// <item><term>&amp;nn</term><description>Change font size. nn is an integer. ex &amp;24</description></item>
        /// <item><term>&amp;G</term><description>Placeholder for images. Images cannot be added by the library, but its possible to use in a template.</description></item>
        /// </list>
        /// </remarks>
        public ExcelHeaderFooter HeaderFooter
        {
            get
            {
                if (_headerFooter == null)
                {
                    XmlNode headerFooterNode = TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager);
                    if (headerFooterNode == null)
                        headerFooterNode = CreateNode("d:headerFooter");
                    _headerFooter = new ExcelHeaderFooter(NameSpaceManager, headerFooterNode, this);
                }
                return (_headerFooter);
            }
        }
        #endregion

        #region "PrinterSettings"
        /// <summary>
        /// Printer settings
        /// </summary>
        public ExcelPrinterSettings PrinterSettings
        {
            get
            {
                var ps = new ExcelPrinterSettings(NameSpaceManager, TopNode, this);
                ps.SchemaNodeOrder = SchemaNodeOrder;
                return ps;
            }
        }
        #endregion

        #endregion // END Worksheet Public Properties
        ExcelSlicerXmlSources _slicerXmlSources = null;
        internal ExcelSlicerXmlSources SlicerXmlSources
        {
            get
            {
                if (_slicerXmlSources == null)
                {
                    _slicerXmlSources = new ExcelSlicerXmlSources(NameSpaceManager, TopNode, Part);
                }
                return _slicerXmlSources;
            }
        }

        #region Worksheet Public Methods

        ///// <summary>
        ///// Provides access to an individual cell within the worksheet.
        ///// </summary>
        ///// <param name="row">The row number in the worksheet</param>
        ///// <param name="col">The column number in the worksheet</param>
        ///// <returns></returns>		
        //internal ExcelCell Cell(int row, int col)
        //{
        //    return new ExcelCell(_values, row, col);
        //}
        /// <summary>
        /// Provides access to a range of cells
        /// </summary>  
        public ExcelRange Cells
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return new ExcelRange(this, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
        }
        /// <summary>
        /// Provides access to the selected range of cells
        /// </summary>  
        public ExcelRange SelectedRange
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return new ExcelRange(this, View.SelectedRange);
            }
        }
        internal MergeCellsCollection _mergedCells = new MergeCellsCollection();
        /// <summary>
        /// Addresses to merged ranges
        /// </summary>
        public MergeCellsCollection MergedCells
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                return _mergedCells;
            }
        }
        /// <summary>
		/// Provides access to an individual row within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <returns></returns>
		public ExcelRow Row(int row)
        {
            CheckSheetTypeAndNotDisposed();
            if (row < 1 || row > ExcelPackage.MaxRows)
            {
                throw (new ArgumentException("Row number out of bounds"));
            }
            return new ExcelRow(this, row);
            //return r;
        }
        /// <summary>
        /// Provides access to an individual column within the worksheet so you can set its properties.
        /// </summary>
        /// <param name="col">The column number in the worksheet</param>
        /// <returns></returns>
        public ExcelColumn Column(int col)
        {
            CheckSheetTypeAndNotDisposed();
            if (col < 1 || col > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("Column number out of bounds"));
            }
            var column = GetValueInner(0, col) as ExcelColumn;
            if (column != null)
            {

                if (column.ColumnMin != column.ColumnMax)
                {
                    int maxCol = column.ColumnMax;
                    column.ColumnMax = col;
                    ExcelColumn copy = CopyColumn(column, col + 1, maxCol);
                }
            }
            else
            {
                int r = 0, c = col;
                if (_values.PrevCell(ref r, ref c))
                {
                    column = GetValueInner(0, c) as ExcelColumn;
                    int maxCol = column.ColumnMax;
                    if (maxCol >= col)
                    {
                        column.ColumnMax = col - 1;
                        if (maxCol > col)
                        {
                            ExcelColumn newC = CopyColumn(column, col + 1, maxCol);
                        }
                        return CopyColumn(column, col, col);
                    }
                }

                column = new ExcelColumn(this, col);
                SetValueInner(0, col, column);
            }
            return column;
        }

        /// <summary>
        /// Returns the name of the worksheet
        /// </summary>
        /// <returns>The name of the worksheet</returns>
        public override string ToString()
        {
            return Name;
        }
        internal ExcelColumn CopyColumn(ExcelColumn c, int col, int maxCol)
        {
            ExcelColumn newC = new ExcelColumn(this, col);
            newC.ColumnMax = maxCol < ExcelPackage.MaxColumns ? maxCol : ExcelPackage.MaxColumns;
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;

            newC.OutlineLevel = c.OutlineLevel;
            newC.Phonetic = c.Phonetic;
            newC.BestFit = c.BestFit;
            newC._width = c._width;
            newC._hidden = c._hidden;
            SetValueInner(0, col, newC);
            return newC;
        }
        /// <summary>
        /// Make the current worksheet active.
        /// </summary>
        public void Select()
        {
            View.TabSelected = true;
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">An address range</param>
        public void Select(string Address)
        {
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// </summary>
        /// <param name="Address">A range of cells</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(string Address, bool SelectSheet)
        {
            CheckSheetTypeAndNotDisposed();
            int fromCol, fromRow, toCol, toRow;
            //Get rows and columns and validate as well
            ExcelCellBase.GetRowColFromAddress(Address, out fromRow, out fromCol, out toRow, out toCol);

            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            View.SelectedRange = Address;
            View.ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">An address range</param>
        public void Select(ExcelAddress Address)
        {
            CheckSheetTypeAndNotDisposed();
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// </summary>
        /// <param name="Address">A range of cells</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(ExcelAddress Address, bool SelectSheet)
        {

            CheckSheetTypeAndNotDisposed();
            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            string selAddress = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column) + ":" + ExcelCellBase.GetAddress(Address.End.Row, Address.End.Column);
            if (Address.Addresses != null)
            {
                foreach (var a in Address.Addresses)
                {
                    selAddress += " " + ExcelCellBase.GetAddress(a.Start.Row, a.Start.Column) + ":" + ExcelCellBase.GetAddress(a.End.Row, a.End.Column);
                }
            }
            View.SelectedRange = selAddress;
            View.ActiveCell = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column);
        }

        #region InsertRow
        /// <summary>
        /// Inserts new rows into the spreadsheet.  Existing rows below the position are 
        /// shifted down.  All formula are updated to take account of the new row(s).
        /// </summary>
        /// <param name="rowFrom">The position of the new row(s)</param>
        /// <param name="rows">Number of rows to insert</param>
        public void InsertRow(int rowFrom, int rows)
        {
            InsertRow(rowFrom, rows, 0);
        }
        /// <summary>
		/// Inserts new rows into the spreadsheet.  Existing rows below the position are 
		/// shifted down.  All formula are updated to take account of the new row(s).
		/// </summary>
        /// <param name="rowFrom">The position of the new row(s)</param>
        /// <param name="rows">Number of rows to insert.</param>
        /// <param name="copyStylesFromRow">Copy Styles from this row. Applied to all inserted rows</param>
		public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
        {
            WorksheetRangeInsertHelper.InsertRow(this, rowFrom, rows, copyStylesFromRow);
        }
        /// <summary>
        /// Inserts new columns into the spreadsheet.  Existing columns below the position are 
        /// shifted down.  All formula are updated to take account of the new column(s).
        /// </summary>
        /// <param name="columnFrom">The position of the new column(s)</param>
        /// <param name="columns">Number of columns to insert</param>        
        public void InsertColumn(int columnFrom, int columns)
        {
            InsertColumn(columnFrom, columns, 0);
        }
        ///<summary>
        /// Inserts new columns into the spreadsheet.  Existing column to the left are 
        /// shifted.  All formula are updated to take account of the new column(s).
        /// </summary>
        /// <param name="columnFrom">The position of the new column(s)</param>
        /// <param name="columns">Number of columns to insert.</param>
        /// <param name="copyStylesFromColumn">Copy Styles from this column. Applied to all inserted columns</param>
        public void InsertColumn(int columnFrom, int columns, int copyStylesFromColumn)
        {
            WorksheetRangeInsertHelper.InsertColumn(this, columnFrom, columns, copyStylesFromColumn);
        }
        #endregion
        #region DeleteRow
        /// <summary>
        /// Delete the specified row from the worksheet.
        /// </summary>
        /// <param name="row">A row to be deleted</param>
        public void DeleteRow(int row)
        {
            DeleteRow(row, 1);
        }
        /// <summary>
        /// Delete the specified rows from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The start row</param>
        /// <param name="rows">Number of rows to delete</param>
        public void DeleteRow(int rowFrom, int rows)
        {
            WorksheetRangeDeleteHelper.DeleteRow(this, rowFrom, rows);
        }

        /// <summary>
        /// Deletes the specified rows from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The number of the start row to be deleted</param>
        /// <param name="rows">Number of rows to delete</param>
        /// <param name="shiftOtherRowsUp">Not used. Rows are always shifted</param>
        [Obsolete("Use the two-parameter method instead")]
        public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
        {
            DeleteRow(rowFrom, rows);
        }
        #endregion
        #region Delete column
        /// <summary>
        /// Delete the specified column from the worksheet.
        /// </summary>
        /// <param name="column">The column to be deleted</param>
        public void DeleteColumn(int column)
        {
            DeleteColumn(column, 1);
        }
        /// <summary>
        /// Delete the specified columns from the worksheet.
        /// </summary>
        /// <param name="columnFrom">The start column</param>
        /// <param name="columns">Number of columns to delete</param>
        public void DeleteColumn(int columnFrom, int columns)
        {
            WorksheetRangeDeleteHelper.DeleteColumn(this, columnFrom, columns);
        }
        #endregion
        /// <summary>
        /// Get the cell value from thw worksheet
        /// </summary>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The row number</param>
        /// <returns>The value</returns>
        public object GetValue(int Row, int Column)
        {
            CheckSheetTypeAndNotDisposed();
            var v = GetValueInner(Row, Column);
            if (v != null)
            {
                //var cell = ((ExcelCell)_cells[cellID]);
                if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
                {
                    return (object)Cells[Row, Column].RichText.Text;
                }
                else
                {
                    return v;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get a strongly typed cell value from the worksheet
        /// </summary>
        /// <typeparam name="T">The type</typeparam>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The row number</param>
        /// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
        public T GetValue<T>(int Row, int Column)
        {
            CheckSheetTypeAndNotDisposed();
            //ulong cellID=ExcelCellBase.GetCellID(SheetID, Row, Column);
            var v = GetValueInner(Row, Column);
            if (v == null)
            {
                return default(T);
            }

            //var cell=((ExcelCell)_cells[cellID]);
            if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
            {
                return (T)(object)Cells[Row, Column].RichText.Text;
            }

            return ConvertUtil.GetTypedCellValue<T>(v);
        }

        /// <summary>
        /// Set the value of a cell
        /// </summary>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The column number</param>
        /// <param name="Value">The value</param>
        public void SetValue(int Row, int Column, object Value)
        {
            CheckSheetTypeAndNotDisposed();
            if (Row < 1 || Column < 1 || Row > ExcelPackage.MaxRows && Column > ExcelPackage.MaxColumns)
            {
                throw new ArgumentOutOfRangeException("Row or Column out of range");
            }
            SetValueInner(Row, Column, Value);
        }
        /// <summary>
        /// Set the value of a cell
        /// </summary>
        /// <param name="Address">The Excel address</param>
        /// <param name="Value">The value</param>
        public void SetValue(string Address, object Value)
        {
            CheckSheetTypeAndNotDisposed();
            int row, col;
            ExcelAddressBase.GetRowCol(Address, out row, out col, true);
            if (row < 1 || col < 1 || row > ExcelPackage.MaxRows && col > ExcelPackage.MaxColumns)
            {
                throw new ArgumentOutOfRangeException("Address is invalid or out of range");
            }
            SetValueInner(row, col, Value);
        }

        #region MergeCellId

        /// <summary>
        /// Get MergeCell Index No
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public int GetMergeCellId(int row, int column)
        {
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(_mergedCells[i]))
                {
                    ExcelRange range = Cells[_mergedCells[i]];

                    if (range.Start.Row <= row && row <= range.End.Row)
                    {
                        if (range.Start.Column <= column && column <= range.End.Column)
                        {
                            return i + 1;
                        }
                    }
                }
            }
            return 0;
        }
        #endregion
        #endregion //End Worksheet Public Methods
        #region Worksheet Private Methods
        internal void UpdateSheetNameInFormulas(string newName, int rowFrom, int rows, int columnFrom, int columns)
        {
            lock (this)
            {
                foreach (var f in _sharedFormulas.Values)
                {
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, columns, rowFrom, columnFrom, Name, newName);
                }
                using (var cse = new CellStoreEnumerator<object>(_formulas))
                {
                    while (cse.Next())
                    {
                        if (cse.Value is string)
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, columns, rowFrom, columnFrom, Name, newName);
                        }
                    }
                }
            }
        }

        private void UpdateSheetNameInFormulas(string oldName, string newName)
        {
            if (string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName))
                throw new ArgumentNullException("Sheet name can't be empty");

            lock (this)
            {
                foreach (var sf in _sharedFormulas.Values)
                {
                    sf.Formula = ExcelCellBase.UpdateSheetNameInFormula(sf.Formula, oldName, newName);
                }
                using (var cse = new CellStoreEnumerator<object>(_formulas))
                {
                    while (cse.Next())
                    {
                        if (cse.Value is string v) //Non shared Formulas 
                        {
                            cse.Value = ExcelCellBase.UpdateSheetNameInFormula(v, oldName, newName);
                        }
                    }
                }
            }
        }
        #region Worksheet Save
        internal void Save()
        {
            DeletePrinterSettings();

            if (_worksheetXml != null)
            {
                SaveDrawings();
                if (!(this is ExcelChartsheet))
                {
                    // save the header & footer (if defined)
                    if (_headerFooter != null)
                        HeaderFooter.Save();

                    var d = Dimension;
                    if (d == null)
                    {
                        this.DeleteAllNode("d:dimension/@ref");
                    }
                    else
                    {
                        this.SetXmlNodeString("d:dimension/@ref", d.Address);
                    }


                    if (Drawings.Count == 0)
                    {
                        //Remove node if no drawings exists.
                        DeleteNode("d:drawing");
                    }

                    SaveVmlDrawings();
                    SaveComments();
                    SaveThreadedComments();
                    HeaderFooter.SaveHeaderFooterImages();
                    SaveTables();
                    if (HasLoadedPivotTables) SavePivotTables();
                    SaveSlicers();
                }
            }
        }

        private void SaveDrawings()
        {
            if (Drawings.UriDrawing != null)
            {
                if (Drawings.Count == 0)
                {
                    Part.DeleteRelationship(Drawings._drawingRelation.Id);
                    _package.ZipPackage.DeletePart(Drawings.UriDrawing);
                }
                else
                {
                    RowHeightCache = new Dictionary<int, double>();
                    foreach (ExcelDrawing d in Drawings)
                    {
                        d.AdjustPositionAndSize();
                        d.UpdatePositionAndSizeXml();
                        HandleSaveForIndividualDrawings(d);
                    }
                    Packaging.ZipPackagePart partPack = Drawings.Part;
                    var stream = partPack.GetStream(FileMode.Create, FileAccess.Write);
                    var xr = new XmlTextWriter(stream, Encoding.UTF8);
                    xr.Formatting = Formatting.None;

                    Drawings.DrawingXml.Save(xr);
                }
            }
        }

        private static void HandleSaveForIndividualDrawings(ExcelDrawing d)
        {
            if (d is ExcelChart c)
            {
                var xr = new XmlTextWriter(c.Part.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
                xr.Formatting = Formatting.None;
                c.ChartXml.Save(xr);
            }
            else if (d is ExcelSlicer<ExcelTableSlicerCache> s)
            {
                s.Cache.SlicerCacheXml.Save(s.Cache.Part.GetStream(FileMode.Create, FileAccess.Write));
            }
            else if (d is ExcelSlicer<ExcelPivotTableSlicerCache> p)
            {
                if (p.Cache == null) return;
                p.Cache.UpdateItemsXml();
                p.Cache.SlicerCacheXml.Save(p.Cache.Part.GetStream(FileMode.Create, FileAccess.Write));
            }
            else if (d is ExcelControl ctrl)
            {
                ctrl.ControlPropertiesXml.Save(ctrl.ControlPropertiesPart.GetStream(FileMode.Create, FileAccess.Write));
                ctrl.UpdateXml();
            }
            if (d is ExcelGroupShape grp)
            {
                foreach (var sd in grp.Drawings)
                {
                    HandleSaveForIndividualDrawings(sd);
                }
            }
        }

        private void SaveSlicers()
        {
            SlicerXmlSources.Save();
        }
        private void SaveThreadedComments()
        {
            if (ThreadedComments != null && ThreadedComments.Threads != null)
            {
                if (!ThreadedComments.Threads.Any(x => x.Comments.Count > 0) && _package.ZipPackage.PartExists(ThreadedCommentsUri))
                {
                    _package.ZipPackage.DeletePart(ThreadedCommentsUri);
                }
                else if (ThreadedComments.Threads.Count() > 0)
                {
                    if (!_package.ZipPackage.PartExists(ThreadedCommentsUri))
                    {
                        var tcUri = ThreadedCommentsUri;
                        _package.ZipPackage.CreatePart(tcUri, "application/vnd.ms-excel.threadedcomments+xml");
                        Part.CreateRelationship(tcUri, Packaging.TargetMode.Internal, ExcelPackage.schemaThreadedComment);
                    }
                    _package.SavePart(ThreadedCommentsUri, ThreadedComments.ThreadedCommentsXml);
                }
            }
        }

        internal void SaveHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            //Init Zip
            stream.CodecBufferSize = 8096;
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            stream.PutNextEntry(fileName);


            SaveXml(stream);
        }


        /// <summary>
        /// Delete the printersettings relationship and part.
        /// </summary>
        private void DeletePrinterSettings()
        {
            //Delete the relationship from the pageSetup tag
            XmlAttribute attr = (XmlAttribute)WorksheetXml.SelectSingleNode("//d:pageSetup/@r:id", NameSpaceManager);
            if (attr != null)
            {
                string relID = attr.Value;
                //First delete the attribute from the XML
                attr.OwnerElement.Attributes.Remove(attr);
                if (Part.RelationshipExists(relID))
                {
                    var rel = Part.GetRelationship(relID);
                    Uri printerSettingsUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    Part.DeleteRelationship(rel.Id);

                    //Delete the part from the package
                    if (_package.ZipPackage.PartExists(printerSettingsUri))
                    {
                        _package.ZipPackage.DeletePart(printerSettingsUri);
                    }
                }
            }
        }
        private void SaveComments()
        {
            if (_comments != null)
            {
                if (_comments.Count == 0)
                {
                    if (_comments.Uri != null)
                    {
                        Part.DeleteRelationship(_comments.RelId);
                        if (_package.ZipPackage.PartExists(_comments.Uri))
                        {
                            _package.ZipPackage.DeletePart(_comments.Uri);
                        }
                    }
                    if (VmlDrawings.Count == 0)
                    {
                        RemoveLegacyDrawingRel(VmlDrawings.RelId);
                    }
                }
                else
                {
                    if (_comments.Uri == null)
                    {
                        var id = SheetId;
                        _comments.Uri = XmlHelper.GetNewUri(_package.ZipPackage, @"/xl/comments{0}.xml", ref id); //Issue 236-Part already exists fix
                    }
                    if (_comments.Part == null)
                    {
                        _comments.Part = _package.ZipPackage.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _package.Compression);
                        var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");
                    }
                    _comments.CommentXml.Save(_comments.Part.GetStream(FileMode.Create));
                }
            }
        }

        private void SaveVmlDrawings()
        {
            if (_vmlDrawings != null)
            {
                if (_vmlDrawings.Count == 0)
                {
                    if (_vmlDrawings.Part != null)
                    {
                        Part.DeleteRelationship(_vmlDrawings.RelId);
                        if (_package.ZipPackage.PartExists(_vmlDrawings.Uri))
                        {
                            _package.ZipPackage.DeletePart(_vmlDrawings.Uri);
                        }
                        DeleteNode($"d:legacyDrawing[@r:id='{_vmlDrawings.RelId}']");
                    }
                }
                else
                {
                    _vmlDrawings.CreateVmlPart();
                }
            }
        }
        /// <summary>
        /// Save all table data
        /// </summary>
        private void SaveTables()
        {
            foreach (var tbl in Tables)
            {
                if (tbl.ShowFilter)
                {
                    tbl.AutoFilter.Save();
                }
                if (tbl.ShowHeader || tbl.ShowTotal)
                {
                    int colNum = tbl.Address._fromCol;
                    var colVal = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
                    foreach (var col in tbl.Columns)
                    {
                        string n = col.Name.ToLowerInvariant();
                        if (tbl.ShowHeader)
                        {
                            var v = tbl.WorkSheet.GetValue(tbl.Address._fromRow, colNum);

                            if (v is string s)
                            {
                                n = s;
                            }
                            else
                            {
                                //Column headers must be a string. Set the value to the string value of the number.
                                n = tbl.WorkSheet.Cells[tbl.Address._fromRow, tbl.Address._fromCol + col.Position].Text;
                                SetValueInner(tbl.Address._fromRow, colNum, n);
                            }

                            if (string.IsNullOrEmpty(n))
                            {
                                n = col.Name.ToLowerInvariant();
                                SetValueInner(tbl.Address._fromRow, colNum, ConvertUtil.ExcelDecodeString(col.Name));
                            }
                            else if (col.Name != n)
                            {
                                col.Name = n;
                                SetValueInner(tbl.Address._fromRow, colNum, ConvertUtil.ExcelDecodeString(col.Name));
                                if (tbl.WorkSheet.IsRichText(tbl.Address._fromRow, colNum))
                                {
                                    _flags.SetFlagValue(tbl.Address._fromRow, colNum, false, CellFlags.RichText);
                                }
                            }
                        }
                        else
                        {
                            n = col.Name.ToLowerInvariant();
                        }

                        if (colVal.Contains(n))
                        {
                            throw (new InvalidDataException(string.Format("Table {0} Column {1} does not have a unique name.", tbl.Name, col.Name)));
                        }
                        colVal.Add(n);
                        colNum++;
                    }
                }
                if (tbl.Part == null)
                {
                    var id = tbl.Id;
                    tbl.TableUri = GetNewUri(_package.ZipPackage, @"/xl/tables/table{0}.xml", ref id);
                    tbl.Id = id;
                    tbl.Part = _package.ZipPackage.CreatePart(tbl.TableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", Workbook._package.Compression);
                    var stream = tbl.Part.GetStream(FileMode.Create);
                    tbl.TableXml.Save(stream);
                    var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, tbl.TableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");
                    tbl.RelationshipID = rel.Id;

                    CreateNode("d:tableParts");
                    XmlNode tbls = TopNode.SelectSingleNode("d:tableParts", NameSpaceManager);

                    var tblNode = tbls.OwnerDocument.CreateElement("tablePart", ExcelPackage.schemaMain);
                    tbls.AppendChild(tblNode);
                    tblNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
                }
                else
                {
                    var stream = tbl.Part.GetStream(FileMode.Create);
                    tbl.TableXml.Save(stream);
                }
            }
        }

        internal bool IsRichText(int row, int col)
        {
            return _flags.GetFlagValue(row, col, CellFlags.RichText);
        }

        internal void SetTableTotalFunction(ExcelTable tbl, ExcelTableColumn col, int colNum = -1)
        {
            if (tbl.ShowTotal == false) return;
            if (colNum == -1)
            {
                for (int i = 0; i < tbl.Columns.Count; i++)
                {
                    if (tbl.Columns[i].Name == col.Name)
                    {
                        colNum = tbl.Address._fromCol + i;
                    }
                }
            }
            if (colNum == -1) return;
            if (col.TotalsRowFunction == RowFunctions.Custom)
            {
                SetFormula(tbl.Address._toRow, colNum, col.TotalsRowFormula);
            }
            else if (col.TotalsRowFunction != RowFunctions.None)
            {
                switch (col.TotalsRowFunction)
                {
                    case RowFunctions.Average:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "101"));
                        break;
                    case RowFunctions.CountNums:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
                        break;
                    case RowFunctions.Count:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "103"));
                        break;
                    case RowFunctions.Max:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "104"));
                        break;
                    case RowFunctions.Min:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "105"));
                        break;
                    case RowFunctions.StdDev:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "107"));
                        break;
                    case RowFunctions.Var:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "110"));
                        break;
                    case RowFunctions.Sum:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "109"));
                        break;
                    default:
                        throw (new Exception("Unknown RowFunction enum"));
                }
            }
            else
            {
                SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);
            }
        }

        internal void SetFormula(int row, int col, object value)
        {
            _formulas.SetValue(row, col, value);
            if (!ExistsValueInner(row, col)) SetValueInner(row, col, null);
        }

        private void SavePivotTables()
        {
            foreach (var pt in PivotTables)
            {
                pt.Save();
            }
        }

        private static string GetTotalFunction(ExcelTableColumn col, string funcNum)
        {
            var escapedName = col.Name.Replace("'", "''");
            escapedName = escapedName.Replace("[", "'[");
            escapedName = escapedName.Replace("]", "']");
            escapedName = escapedName.Replace("#", "'#");
            return string.Format("SUBTOTAL({0},{1}[{2}])", funcNum, col._tbl.Name, escapedName);
        }

        private void SaveXml(Stream stream)
        {
            //Create the nodes if they do not exist.
            StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.UTF8, 65536);
            if (this is ExcelChartsheet)
            {
                sw.Write(_worksheetXml.OuterXml);
            }
            else
            {

                if (_autoFilter != null)
                {
                    _autoFilter.Save();
                }
                CreateNode("d:cols");
                CreateNode("d:sheetData");
                CreateNode("d:mergeCells");
                CreateNode("d:hyperlinks");
                CreateNode("d:rowBreaks");
                CreateNode("d:colBreaks");

                if (DataValidations != null && DataValidations.Count != 0)
                {
                    WorksheetXml.DocumentElement.SetAttribute("xmlns:xr", ExcelPackage.schemaXr);
                    WorksheetXml.DocumentElement.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);

                    var ignorables = WorksheetXml.DocumentElement.GetAttribute("Ignorable", ExcelPackage.schemaMarkupCompatibility);
                    if (ignorables != null)
                    {
                        var namespaces = ignorables.Split(' ');
                        if(!namespaces.Any(x => x == "xr"))
                        {
                            WorksheetXml.DocumentElement.SetAttribute("Ignorable", ExcelPackage.schemaMarkupCompatibility, ignorables + " xr");
                        }
                    }
                    else
                    {
                        WorksheetXml.DocumentElement.SetAttributeNode("Ignorable", ExcelPackage.schemaMarkupCompatibility);
                        WorksheetXml.DocumentElement.SetAttribute("Ignorable", ExcelPackage.schemaMarkupCompatibility, "xr");
                    }

                    if (DataValidations.HasValidationType(InternalValidationType.DataValidation))
                    {
                        var node = (XmlElement)CreateNode("d:dataValidations");

                        if (DataValidations.HasValidationType(InternalValidationType.ExtLst) &&
                            GetNode("d:extLst") == null)
                        {
                            CreateNode("d:extLst");
                        }
                    }
                    else if (GetNode("d:extLst") == null)
                    {
                        CreateNode("d:extLst");
                    }
                }

                var prefix = GetNameSpacePrefix();
                var xml = _worksheetXml.OuterXml;
                int startOfNode = 0, endOfNode = 0;

                ExcelXmlWriter writer = new ExcelXmlWriter(this, _package);
                writer.WriteNodes(sw, xml, ref startOfNode, ref endOfNode);
            }
            sw.Flush();
        }

        internal string GetNameSpacePrefix()
        {
            if (_worksheetXml.DocumentElement == null) return "";
            foreach (XmlAttribute a in _worksheetXml.DocumentElement.Attributes)
            {
                if (a.Value == ExcelPackage.schemaMain)
                {
                    if (string.IsNullOrEmpty(a.Prefix))
                    {
                        return "";
                    }
                    else
                    {
                        return a.LocalName + ":";
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// Dimension address for the worksheet. 
        /// Top left cell to Bottom right.
        /// If the worksheet has no cells, null is returned
        /// </summary>
        public ExcelAddressBase Dimension
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                int fromRow, fromCol, toRow, toCol;
                if (_values.GetDimension(out fromRow, out fromCol, out toRow, out toCol))
                {
                    var addr = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
                    addr._ws = Name;
                    return addr;
                }
                else
                {
                    return null;
                }
            }
        }
        ExcelSheetProtection _protection = null;
        /// <summary>
        /// Access to sheet protection properties
        /// </summary>
        public ExcelSheetProtection Protection
        {
            get
            {
                if (_protection == null)
                {
                    _protection = new ExcelSheetProtection(NameSpaceManager, TopNode, this);
                }
                return _protection;
            }
        }

        private ExcelProtectedRangeCollection _protectedRanges = null;
        /// <summary>
        /// Access to protected ranges in the worksheet
        /// </summary>
        public ExcelProtectedRangeCollection ProtectedRanges
        {
            get
            {
                if (_protectedRanges == null)
                {
                    _protectedRanges = new ExcelProtectedRangeCollection(this);
                }
                return _protectedRanges;
            }
        }

        #region Drawing
        internal bool HasDrawingRelationship
        {
            get
            {
                return WorksheetXml.DocumentElement.SelectSingleNode("d:drawing", NameSpaceManager) != null;
            }
        }

        internal Dictionary<int, double> RowHeightCache { get; set; } = new Dictionary<int, double>();
        internal ExcelDrawings _drawings = null;
        /// <summary>
        /// Collection of drawing-objects like shapes, images and charts
        /// </summary>
        public ExcelDrawings Drawings
        {
            get
            {
                LoadDrawings();
                return _drawings;
            }
        }

        internal void LoadDrawings()
        {
            if (_drawings == null)
            {
                _drawings = new ExcelDrawings(_package, this);
            }
        }
        #endregion
        #region SparklineGroups
        ExcelSparklineGroupCollection _sparklineGroups = null;
        /// <summary>
        /// Collection of Sparkline-objects. 
        /// Sparklines are small in-cell charts.
        /// </summary>
        public ExcelSparklineGroupCollection SparklineGroups
        {
            get
            {
                if (_sparklineGroups == null)
                {
                    _sparklineGroups = new ExcelSparklineGroupCollection(this);
                }
                return _sparklineGroups;
            }
        }
        #endregion
        ExcelTableCollection _tables = null;
        /// <summary>
        /// Tables defined in the worksheet.
        /// </summary>
        public ExcelTableCollection Tables
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (Workbook._nextTableID == int.MinValue) Workbook.ReadAllTables();
                if (_tables == null)
                {
                    _tables = new ExcelTableCollection(this);
                }
                return _tables;
            }
        }
        internal ExcelPivotTableCollection _pivotTables = null;
        /// <summary>
        /// Pivottables defined in the worksheet.
        /// </summary>
        public ExcelPivotTableCollection PivotTables
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (_pivotTables == null)
                {
                    _pivotTables = new ExcelPivotTableCollection(this);
                    if (Workbook._nextPivotTableID == int.MinValue) Workbook.ReadAllPivotTables();
                }
                return _pivotTables;
            }
        }
        internal bool HasLoadedPivotTables
        {
            get
            {
                return _pivotTables != null;
            }
        }
        private ExcelConditionalFormattingCollection _conditionalFormatting = null;
        /// <summary>
        /// ConditionalFormatting defined in the worksheet. Use the Add methods to create ConditionalFormatting and add them to the worksheet. Then
        /// set the properties on the instance returned.
        /// </summary>
        /// <seealso cref="ExcelConditionalFormattingCollection"/>
        public ExcelConditionalFormattingCollection ConditionalFormatting
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (_conditionalFormatting == null)
                {
                    _conditionalFormatting = new ExcelConditionalFormattingCollection(this);
                }
                return _conditionalFormatting;
            }
        }

        ExcelIgnoredErrorCollection _ignoredErrors = null;
        /// <summary>
        /// Ignore Errors for the specified ranges and error types.
        /// </summary>
        public ExcelIgnoredErrorCollection IgnoredErrors
        {
            get
            {
                CheckSheetTypeAndNotDisposed();
                if (_ignoredErrors == null)
                {
                    _ignoredErrors = new ExcelIgnoredErrorCollection(_package, this, NameSpaceManager);
                }
                return _ignoredErrors;
            }
        }
        internal void ClearValidations()
        {
            _dataValidations = null;
        }

        ExcelBackgroundImage _backgroundImage = null;
        /// <summary>
        /// An image displayed as the background of the worksheet.
        /// </summary>
        public ExcelBackgroundImage BackgroundImage
        {
            get
            {
                if (_backgroundImage == null)
                {
                    _backgroundImage = new ExcelBackgroundImage(NameSpaceManager, TopNode, this);
                }
                return _backgroundImage;
            }
        }
        /// <summary>
        /// The workbook object
        /// </summary>
        public ExcelWorkbook Workbook
        {
            get
            {
                return _package.Workbook;
            }
        }

        #endregion
        #endregion  // END <Worksheet Private Methods

        /// <summary>
        /// Get the next ID from a shared formula or an Array formula
        /// Sharedforumlas will have an id from 0-x. Array formula ids start from 0x4000001-. 
        /// </summary>
        /// <param name="isArray">If the formula is an array formula</param>
        /// <returns></returns>
        internal int GetMaxShareFunctionIndex(bool isArray)
        {
            int i = _sharedFormulas.Count + 1;
            if (isArray)
                i |= 0x40000000;

            while (_sharedFormulas.ContainsKey(i))
            {
                i++;
            }
            return i;
        }
        internal void SetHFLegacyDrawingRel(string relID)
        {
            SetXmlNodeString("d:legacyDrawingHF/@r:id", relID);
        }
        internal void RemoveLegacyDrawingRel(string relID)
        {
            var n = WorksheetXml.DocumentElement.SelectSingleNode(string.Format("d:legacyDrawing[@r:id=\"{0}\"]", relID), NameSpaceManager);
            if (n != null)
            {
                n.ParentNode.RemoveChild(n);
            }
        }

        internal void UpdateCellsWithDate1904Setting()
        {
            var cse = new CellStoreEnumerator<ExcelValue>(_values);
            var offset = Workbook.Date1904 ? -ExcelWorkbook.date1904Offset : ExcelWorkbook.date1904Offset;
            while (cse.MoveNext())
            {
                if (cse.Value._value is DateTime)
                {
                    try
                    {
                        double sdv = ((DateTime)cse.Value._value).ToOADate();
                        sdv += offset;

                        //cse.Value._value = DateTime.FromOADate(sdv);
                        SetValueInner(cse.Row, cse.Column, DateTime.FromOADate(sdv));
                    }
                    catch
                    {
                    }
                }
            }
        }
        internal string GetFormula(int row, int col)
        {
            var v = _formulas?.GetValue(row, col);
            if (v is int)
            {
                return _sharedFormulas[(int)v].GetFormula(row, col, Name);
            }
            else if (v != null)
            {
                return v.ToString();
            }
            else
            {
                return "";
            }
        }
        internal string GetFormulaR1C1(int row, int col)
        {
            var v = _formulas?.GetValue(row, col);
            if (v is int)
            {
                var sf = _sharedFormulas[(int)v];
                return R1C1Translator.ToR1C1Formula(sf.Formula, sf.StartRow, sf.StartCol);
            }
            else if (v != null)
            {
                return R1C1Translator.ToR1C1Formula(v.ToString(), row, col);
            }
            else
            {
                return "";
            }
        }

        private void DisposeInternal(IDisposable candidateDisposable)
        {
            if (candidateDisposable != null)
            {
                candidateDisposable.Dispose();
            }
        }

        /// <summary>
        /// Disposes the worksheet
        /// </summary>
        public void Dispose()
        {
            DisposeInternal(_values);
            DisposeInternal(_formulas);
            DisposeInternal(_flags);
            DisposeInternal(_hyperLinks);
            DisposeInternal(_commentsStore);
            DisposeInternal(_formulaTokens);
            DisposeInternal(_metadataStore);

            _values = null;
            _formulas = null;
            _flags = null;
            _hyperLinks = null;
            _commentsStore = null;
            _formulaTokens = null;
            _metadataStore = null;

            _package = null;
            _pivotTables = null;
            _protection = null;
            if (_sharedFormulas != null) _sharedFormulas.Clear();
            _sharedFormulas = null;
            _sheetView = null;
            _tables = null;
            _vmlDrawings = null;
            _conditionalFormatting = null;
            _dataValidations = null;
            _drawings = null;

            _sheetID = -1;
            _positionId = -1;
        }

        /// <summary>
        /// Get the ExcelColumn for column (span ColumnMin and ColumnMax)
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        internal ExcelColumn GetColumn(int column)
        {
            var c = GetValueInner(0, column) as ExcelColumn;
            if (c == null)
            {
                int row = 0, col = column;
                if (_values.PrevCell(ref row, ref col))
                {
                    c = GetValueInner(0, col) as ExcelColumn;
                    if (c != null && c.ColumnMax >= column)
                    {
                        return c;
                    }
                    return null;
                }
            }
            return c;

        }
        /// <summary>
        /// Check if a worksheet is equal to another
        /// </summary>
        /// <param name="x">First worksheet </param>
        /// <param name="y">Second worksheet</param>
        /// <returns></returns>
        public bool Equals(ExcelWorksheet x, ExcelWorksheet y)
        {
            return x.Name == y.Name && x.SheetId == y.SheetId && x.WorksheetXml.OuterXml == y.WorksheetXml.OuterXml;
        }
        /// <summary>
        /// Returns a hashcode generated from the WorksheetXml
        /// </summary>
        /// <param name="obj">The worksheet</param>
        /// <returns>The hashcode</returns>
        public int GetHashCode(ExcelWorksheet obj)
        {
            return obj.WorksheetXml.OuterXml.GetHashCode();
        }
        ControlsCollectionInternal _controls = null;
        internal ControlsCollectionInternal Controls
        {
            get
            {
                if (_controls == null)
                {
                    _controls = new ControlsCollectionInternal(NameSpaceManager, TopNode);
                }
                return _controls;
            }
        }
        /// <summary>
        /// A collection of row specific properties in the worksheet.
        /// </summary>
        public ExcelRowsCollection Rows
        {
            get
            {
                return new ExcelRowsCollection(this);
            }
        }
        /// <summary>
        /// A collection of column specific properties in the worksheet.
        /// </summary>
        public ExcelColumnCollection Columns
        {
            get
            {
                return new ExcelColumnCollection(this);
            }
        }

        internal bool IsDisposed
        {
            get
            {
                return _values == null;
            }
        }

        ExcelPackage IPictureRelationDocument.Package { get { return _package; } }

        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();
        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes { get { return _hashes; } }

        Packaging.ZipPackagePart IPictureRelationDocument.RelatedPart { get { return Part; } }

        Uri IPictureRelationDocument.RelatedUri { get { return _worksheetUri; } }
        #region Worksheet internal Accessor
        /// <summary>
        /// Get accessor of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>cell value</returns>
        internal ExcelValue GetCoreValueInner(int row, int col)
        {
            return _values.GetValue(row, col);
        }
        /// <summary>
        /// Get accessor of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>cell value</returns>
        internal object GetValueInner(int row, int col)
        {
            return _values.GetValue(row, col)._value;
        }
        /// <summary>
        /// Get accessor of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>cell styleId</returns>
        internal int GetStyleInner(int row, int col)
        {
            return _values.GetValue(row, col)._styleId;
        }

        /// <summary>
        /// Set accessor of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value">value</param>
        internal void SetValueInner(int row, int col, object value)
        {
            _values.SetValue_Value(row, col, value);
        }
        /// <summary>
        /// Set accessor of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="styleId">styleId</param>
        internal void SetStyleInner(int row, int col, int styleId)
        {
            _values.SetValue_Style(row, col, styleId);
        }
        /// <summary>
        /// Set accessor of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value">value</param>
        /// <param name="styleId">styleId</param>
        internal void SetValueStyleIdInner(int row, int col, object value, int styleId)
        {
            _values.SetValue(row, col, value, styleId);
        }
        /// <summary>
        /// Bulk(Range) set accessor of sheet value, for value array
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromColumn">start column</param>
        /// <param name="toRow">end row</param>
        /// <param name="toColumn">end column</param>
        /// <param name="values">set values</param>
        /// <param name="setHyperLinkFromValue">If the value is of type Uri or ExcelHyperlink the Hyperlink property is set.</param>
        internal void SetRangeValueInner(int fromRow, int fromColumn, int toRow, int toColumn, object[,] values, bool setHyperLinkFromValue)
        {
            if (setHyperLinkFromValue)
            {
                SetValuesWithHyperLink(fromRow, fromColumn, values);
            }
            else
            {
                _values.SetValueRange_Value(fromRow, fromColumn, values);
            }
            //Clearout formulas and flags, for example the rich text flag.
            _formulas.Clear(fromRow, fromColumn, values.GetUpperBound(0) + 1, values.GetUpperBound(1) + 1);
            _flags.Clear(fromRow, fromColumn, values.GetUpperBound(0) + 1, values.GetUpperBound(1) + 1);
            _metadataStore.Clear(fromRow, fromColumn, values.GetUpperBound(0) + 1, values.GetUpperBound(1) + 1);
        }

        private void SetValuesWithHyperLink(int fromRow, int fromColumn, object[,] values)
        {
            var rowBound = values.GetUpperBound(0);
            var colBound = values.GetUpperBound(1);

            for (int r = 0; r <= rowBound; r++)
            {
                for (int c = 0; c <= colBound; c++)
                {
                    var v = values[r, c];
                    var row = fromRow + r;
                    var col = fromColumn + c;
                    if (v == null)
                    {
                        _values.SetValue_Value(row, col, v);
                        continue;
                    }
                    var t = v.GetType();
                    if (t == typeof(Uri) || t == typeof(ExcelHyperLink))
                    {
                        _hyperLinks.SetValue(row, col, (Uri)v);
                        if (v is ExcelHyperLink hl)
                        {
                            SetValueInner(row, col, hl.Display);
                        }
                        else
                        {
                            var cv = GetValueInner(row, col);
                            if (cv == null || cv.ToString() == "")
                            {
                                SetValueInner(row, col, ((Uri)v).OriginalString);
                            }
                        }
                    }
                    else
                    {
                        _values.SetValue_Value(row, col, v);
                    }
                }
            }
        }

        /// <summary>
        /// Existance check of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>is exists</returns>
        internal bool ExistsValueInner(int row, int col)
        {
            return (_values.GetValue(row, col)._value != null);
        }
        /// <summary>
        /// Existance check of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>is exists</returns>
        internal bool ExistsStyleInner(int row, int col)
        {
            return (_values.GetValue(row, col)._styleId > 0);
        }
        /// <summary>
        /// Existence check of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value"></param>
        /// <returns>is exists</returns>
        internal bool ExistsValueInner(int row, int col, ref object value)
        {
            value = _values.GetValue(row, col)._value;
            return (value != null);
        }
        /// <summary>
        /// Existence check of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="styleId"></param>
        /// <returns>is exists</returns>
        internal bool ExistsStyleInner(int row, int col, ref int styleId)
        {
            styleId = _values.GetValue(row, col)._styleId;
            return (styleId > 0);
        }
        internal void RemoveSlicerReference(ExcelSlicerXmlSource xmlSource)
        {
            var node = GetNode($"d:extLst/d:ext/x14:slicerList/x14:slicer[@r:id='{xmlSource.Rel.Id}']");
            if (node != null)
            {
                if (node.ParentNode.ChildNodes.Count > 1)
                {
                    node.ParentNode.RemoveChild(node);
                }
                else
                {
                    //Remove the entire ext element.
                    node.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode);
                }
            }
            SlicerXmlSources.Remove(xmlSource);
        }

        internal XmlNode CreateControlContainerNode()
        {
            var node = GetNode("mc:AlternateContent/mc:Choice[@Requires='x14']");
            XmlNode controlsNode;
            if (node == null)
            {
                node = CreateAlternateContentNode("d:controls", "x14");
                controlsNode = node.ChildNodes[0].ChildNodes[0];
            }
            else
            {
                controlsNode = node.SelectSingleNode("d:controls", NameSpaceManager);
                if (controlsNode == null)
                {
                    var f = XmlHelperFactory.Create(NameSpaceManager, node);
                    return f.CreateNode("d:controls");
                }
            }

            var xh = XmlHelperFactory.Create(NameSpaceManager, controlsNode);
            var altNode = (XmlElement)xh.CreateNode("mc:AlternateContent", false, true);

            xh = XmlHelperFactory.Create(NameSpaceManager, altNode);
            var ctrlContainerNode = (XmlElement)xh.CreateNode("mc:Choice");
            ctrlContainerNode.SetAttribute("Requires", "x14");

            return ctrlContainerNode;
        }

        internal void NormalStyleChange()
        {
            _defaultRowHeight = double.NaN;
        }
        #endregion
    }  // END class Worksheet
}
