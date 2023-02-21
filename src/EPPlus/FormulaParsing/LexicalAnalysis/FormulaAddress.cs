using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;
using static OfficeOpenXml.ExcelAddressBase;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    internal class Formula
    {
        internal ExcelWorksheet _ws;
        internal int StartRow, StartCol;
        internal static ISourceCodeTokenizer _tokenizer = OptimizedSourceCodeTokenizer.Default;
        internal IList<Token> Tokens;
        internal IList<Token> RpnTokens;
        internal IExpressionCompiler _compiler;
        internal int AddressExpressionIndex;
        internal CellStoreEnumerator<object> _formulaEnumerator;
        internal ulong _id=ulong.MinValue;
        public ulong Id 
        {
            get
            {
                if(_id== ulong.MinValue)
                {
                    _id=ExcelCellBase.GetCellId(_ws.IndexInList, StartRow, StartCol);
                }
                return _id;
            }
        }

        public Formula()
        {
        }
        public Formula(ExcelWorksheet ws)
        {
            _ws = ws;
        }
        protected void SetFormula(ExcelWorksheet ws, string formula)
        {
            Tokens = _tokenizer.Tokenize(formula);
            RpnTokens = RpnExpressionGraph.CreateRPNTokens(Tokens);
        }
        internal FormulaType FormulaType { get; set; }
        public bool FirstCellDeleted { get; set; }  //del1
        public bool SecondCellDeleted { get; set; } //del2

        public bool DataTableIsTwoDimesional { get; set; } //dt2D
        public bool IsDataTableRow { get; set; } //dtr
        public string R1CellAddress { get; set; } //r1
        public string R2CellAddress { get; set; } //r2

        public Formula(ExcelWorksheet ws, int row, int col)
        {
            _ws = ws;
            StartRow = row;
            StartCol = col;
        }
        public Formula(ExcelWorksheet ws, int row, int col, string formula) : this(ws,row,col)
        {
            SetFormula(ws, formula);            
        }
        internal void SetRowCol(int row, int col)
        {
            StartRow = row;
            StartCol = col;
        }
    }
    internal class SharedFormula : Formula
    {        
        internal int EndRow, EndCol;
        int _rowOffset = 0, _colOffset = 0;
        public SharedFormula() : base()
        {
        }
        public SharedFormula(ExcelWorksheet ws) : base(ws)
        {
        }
        public SharedFormula(ExcelWorksheet ws, string address, string formula) : base(ws)
        {
            _ws = ws;
            Formula = formula;
            ExcelCellBase.GetRowColFromAddress(address, out StartRow, out StartCol, out EndRow, out EndCol);
        }
        public SharedFormula(ExcelRangeBase range) : base(range.Worksheet, range._fromRow, range._fromCol)
        {
            EndRow = range._toRow;
            EndCol = range._toCol;
        }

        public SharedFormula(ExcelRangeBase range, string formula) : this(range.Worksheet, range._fromRow, range._fromCol, range._toRow, range._toCol, formula)
        {
        }

        public SharedFormula(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, string formula) : base(ws, fromRow, fromCol, formula)
        {
            EndRow = toRow;
            EndCol = toCol;
            Formula = formula;
        }
        internal int Index { get; set; }
        string _formula;

        internal string Formula 
        { 
            get
            {
                return _formula;
            }
            set
            {
                _formula = value;
                SetFormula(_ws, value);
            }
        }
        internal bool IsArray { get; set; }
        internal string Address
        {
            get
            {
                return ExcelCellBase.GetAddress(StartRow, StartCol, EndRow, EndCol);
            }
            set
            {
                ExcelCellBase.GetRowColFromAddress(value, out StartRow, out StartCol, out EndRow, out EndCol);
            }
        }
        internal SharedFormula Clone()
        {
            return new SharedFormula(_ws, StartRow, StartCol, EndRow, EndCol, Formula)
            {
                Index = Index,
                FormulaType = FormulaType,
                IsArray = IsArray,
                Tokens = Tokens,
                RpnTokens = RpnTokens,
                Address = Address,
                DataTableIsTwoDimesional = DataTableIsTwoDimesional,
                IsDataTableRow = IsDataTableRow,
                R1CellAddress = R1CellAddress,
                R2CellAddress = R2CellAddress,
                FirstCellDeleted = FirstCellDeleted,
                SecondCellDeleted = SecondCellDeleted,
                _ws = _ws,                
            };
        }
        internal Formula GetFormula(int row, int col)
        {
            return new Formula(_ws, row, col)
            {
                AddressExpressionIndex = 0,
                Tokens = Tokens,
                RpnTokens = RpnTokens,                
                //ExpressionTree = GetExpressionTree(row, col),
                StartRow = row,
                StartCol = col,
                _compiler = _compiler,
            };
        }
        //RpnCompiledFormula _compiledExpressions = null;
        private Dictionary<int, RpnExpression> _compiledExpressions;
        internal RpnFormula GetRpnFormula(RpnOptimizedDependencyChain depChain, int row, int col)
        {
            if (_compiledExpressions == null)
            {
                _compiledExpressions = depChain._graph.CompileExpressions(ref RpnTokens);
            }
            return new RpnFormula(_ws, row, col)
            {
                _tokenIndex = 0,
                _row = row,
                _column = col,
                _tokens = RpnTokens,
                _expressions = CloneExpressions(row, col)
            };
        }
        private Dictionary<int, RpnExpression> CloneExpressions(int row, int col)
        {
            var l=new Dictionary<int, RpnExpression>();
            foreach(var expression in _compiledExpressions)
            {
                if(expression.Value.ExpressionType == ExpressionType.CellAddress ||
                   expression.Value.ExpressionType == ExpressionType.ExcelRange ||
                   expression.Value.ExpressionType == ExpressionType.TableAddress)
                {
                    l.Add(expression.Key, expression.Value.CloneWithOffset(row - StartRow, col - StartCol));
                }
                else
                {
                    l.Add(expression.Key, expression.Value);
                }
            }
            return l;
        }

        internal string GetFormula(int row, int column, string worksheet)
        {
            if (StartRow == row && StartCol == column)
            {
                return Formula;
            }

            SetTokens(worksheet);
            string f = "";
            for(int i=0;i<Tokens.Count;i++)
            {
                var token = Tokens[i];
                if (token.TokenTypeIsSet(TokenType.CellAddress))
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
                else if (token.TokenTypeIsSet(TokenType.FullRowAddress))
                {
                    if (token.Value.StartsWith("$") == false)
                    {
                        if(int.TryParse(token.Value, out int r))
                        {
                            r += row - StartRow;
                            if (r >= 1 && r <= ExcelPackage.MaxRows)
                            {
                                f += r.ToString(CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                f += "#REF!";
                            }
                        }
                        else
                        {
                            f += "#REF!";
                        }
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.FullColumnAddress))
                {
                    if (token.Value.StartsWith("$") == false)
                    {
                        var c = ExcelCellBase.GetColumn(token.Value);
                        c += column - StartCol;
                        if (c >= 1 && c <= ExcelPackage.MaxColumns)
                        { 
                            f += ExcelCellBase.GetColumnLetter(c);
                        }
                        else
                        {
                            f += "#REF!";
                        }
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                else
                {
                    if (token.TokenTypeIsSet(TokenType.StringContent))
                    {
                        f += "\"" + token.Value.Replace("\"", "\"\"") + "\"";
                    }
                    else
                    {
                        f += token.Value;
                    }
                }

            }
            return f;
        }
        internal void SetTokens(string worksheet)
        {
            if (Tokens == null)
            {
                Tokens = _tokenizer.Tokenize(Formula, worksheet);
            }
        }
        //internal Dictionary<ulong, ExpressionTree> _expressionTrees=new Dictionary<ulong, ExpressionTree>();
        //internal override ExpressionTree GetExpressionTree(int row, int col)
        //{
        //    if(row==StartRow && col == StartCol)
        //    {
        //        return ExpressionTree;
        //    }
        //    else
        //    {
        //        var id = ExcelAddressBase.GetCellId(0, row, col);
        //        if(_expressionTrees.TryGetValue(id, out ExpressionTree tree))
        //        {
        //            return tree;
        //        }
        //        else
        //        {
        //            tree= ExpressionTree.CreateFromOffset(row - StartRow, col - StartCol);
        //            _expressionTrees.Add(id, tree);
        //            return tree;
        //        }
        //    }
        //}
    }
    internal enum FormulaType
    {
        Normal,
        Shared,
        Array,
        DataTable
    }
    [Flags]
    internal enum FixedFlag : byte
    {
        None = 0,
        FromRowFixed = 0x1,
        FromColFixed = 0x2,
        ToRowFixed = 0x4,
        ToColFixed = 0x8,
        All = 0xF,
    }
    //internal abstract class TokenInfo
    //{
    //    internal FormulaType Type;
    //    internal short TokenStartPosition;
    //    internal short TokenEndPosition;
    //    internal virtual void SetOffset(int rowOffset, int colOffset) { }

    //    internal abstract string GetValue();

    //    internal virtual bool IsFixed { get { return true; } }
    //}
    //internal class FormulaCellAddress : FormulaAddressBase
    //{
    //    internal FormulaCellAddress()
    //    {
    //    }
    //    internal FormulaCellAddress(FormulaAddressBase addressBase)
    //    {
    //        ExternalReferenceIx = addressBase.ExternalReferenceIx;
    //        WorksheetIx = addressBase.WorksheetIx;
    //    }
    //    internal int Row, Col;
    //    internal bool FixedRow, FixedCol;
    //    //internal override void SetOffset(int rowOffset, int colOffset)
    //    //{
    //    //    if (!FixedRow) Row += rowOffset;
    //    //    if (!FixedCol) Col += colOffset;
    //    //}
    //    //internal override bool IsFixed { get { return FixedRow & FixedCol; } }
    //    //internal override string GetValue()
    //    //{
    //    //    return ExcelCellBase.GetAddress(Row, FixedRow, Col, FixedCol);
    //    //}
    //}
    //internal class FormulaFixedValue : TokenInfo
    //{
    //    public FormulaFixedValue(short startPos, short endPos, object v)
    //    {
    //        Type = FormulaType.FixedValue;
    //        TokenStartPosition = startPos;
    //        TokenEndPosition = endPos;
    //        Value = v;
    //    }
    //    internal object Value;
    //    internal override string GetValue()
    //    {
    //        return Value.ToString();
    //    }
    //}
    //internal class FormulaNamedFormula : TokenInfo
    //{
    //    public FormulaNamedFormula(short startPos, short endPos, string f)
    //    {
    //        Type = FormulaType.Formula;
    //        TokenStartPosition = startPos;
    //        TokenEndPosition = endPos;
    //        Formula = f;
    //    }
    //    internal string Formula;
    //    internal override bool IsFixed { get { return false; } } //TODO: Check here if we can us fixed from the actual formula in  later stage.
    //    internal override string GetValue()
    //    {
    //        return Formula;
    //    }
    //}
    //internal class FormulaRange : TokenInfo
    //{
    //    ParsingContext _ctx;
    //    public FormulaRange(ParsingContext ctx)
    //    {
    //        _ctx = ctx;
    //    }
    //    internal override void SetOffset(int rowOffset, int colOffset)
    //    { 
    //        for(int i=0;i < Ranges.Count;i++)
    //        {
    //            var r=Ranges[i];
    //            if ((r.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.None) r.FromRow += rowOffset;
    //            if ((r.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.None) r.ToRow += rowOffset;
    //            if ((r.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.None) r.FromCol += colOffset;
    //            if ((r.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.None) r.ToCol += colOffset;
    //        }
    //    }
    //    internal override bool IsFixed 
    //    {
    //        get
    //        {
    //            foreach(var r in Ranges)
    //            {
    //                if(r.FixedFlag != FixedFlag.All)
    //                {
    //                    return false;
    //                }
    //            }
    //            return true;
    //        }
    //    }
    //    internal List<FormulaRangeAddress> Ranges;
    //    internal FormulaRange(short startPos, short endPos, int fromRow, int fromCol, int toRow, int toCol, FixedFlag fixedFlag)
    //    {
    //        Type = FormulaType.FormulaRange;
    //        TokenStartPosition = startPos;
    //        TokenEndPosition = endPos;
    //        Ranges = new List<FormulaRangeAddress>();
    //        Ranges.Add(
    //            new FormulaRangeAddress(_ctx)
    //            {
    //                FromRow = fromRow,
    //                FromCol = fromCol,
    //                ToRow = toRow,
    //                ToCol = toCol,
    //                FixedFlag = fixedFlag
    //            });
    //    }
    //    internal FormulaRange(short startPos, short endPos, ExcelRangeBase range)
    //    {
    //        Type = FormulaType.FormulaRange;
    //        TokenStartPosition = startPos;
    //        TokenEndPosition = endPos;
    //        Ranges = new List<FormulaRangeAddress>();
    //        if (range.Addresses == null)
    //        {
    //            Ranges.Add(
    //                new FormulaRangeAddress(_ctx)
    //                {
    //                    ExternalReferenceIx = (short)(string.IsNullOrEmpty(range._wb) ? 0 : range._workbook.ExternalLinks.GetExternalLink(range._wb)),
    //                    WorksheetIx = (short)range.Worksheet.PositionId,
    //                    FromRow = range._fromRow,
    //                    FromCol = range._fromCol,
    //                    ToRow = range._toRow,
    //                    ToCol = range._toCol,

    //                    FixedFlag = (range._fromRowFixed ? FixedFlag.FromRowFixed : 0) |
    //                                (range._fromColFixed ? FixedFlag.FromColFixed : 0) |
    //                                (range._toRowFixed ? FixedFlag.ToRowFixed : 0) |
    //                                (range._toColFixed ? FixedFlag.ToColFixed : 0)
    //                }); 
    //        }
    //        else
    //        {
    //            foreach (var a in range.Addresses)
    //            {
    //                Ranges.Add(
    //                    new FormulaRangeAddress(_ctx)
    //                    {
    //                        ExternalReferenceIx = (short)(string.IsNullOrEmpty(a._wb) ? -1 : range._workbook.ExternalLinks.GetExternalLink(a._wb)),
    //                        WorksheetIx = (short)(string.IsNullOrEmpty(a.WorkSheetName) ? range.Worksheet.PositionId : (range._workbook.Worksheets[a.WorkSheetName]==null ? -1 : range._workbook.Worksheets[a.WorkSheetName].PositionId)),
    //                        FromRow = a._fromRow,
    //                        FromCol = a._fromCol,
    //                        ToRow = a._toRow,
    //                        ToCol = a._toCol,
    //                        FixedFlag = (a._fromRowFixed ? FixedFlag.FromRowFixed : 0) |
    //                                    (a._fromColFixed ? FixedFlag.FromColFixed : 0) |
    //                                    (a._toRowFixed ? FixedFlag.ToRowFixed : 0) |
    //                                    (a._toColFixed ? FixedFlag.ToColFixed : 0) 

    //                    });
    //            }
    //        }
    //    }
    //    internal override string GetValue()
    //    {
    //        var sb=new StringBuilder();
    //        foreach(var r in Ranges)
    //        {
    //            sb.Append(ExcelCellBase.GetAddress(r.FromRow, r.FromCol, r.ToRow, r.ToCol,
    //                (r.FixedFlag & FixedFlag.FromRowFixed) > 0,
    //                (r.FixedFlag & FixedFlag.FromColFixed) > 0,
    //                (r.FixedFlag & FixedFlag.ToRowFixed) > 0,
    //                (r.FixedFlag & FixedFlag.ToColFixed) > 0));
    //            sb.Append(':');
    //        }
    //        return sb.ToString(0, sb.Length - 1);
    //    }
    //}
    public struct FormulaCellAddress
    {
        public FormulaCellAddress(int wsIx, int row, int column)
        {
            WorksheetIx = wsIx;
            Row = row;
            Column = column;
        }
        /// <summary>
        /// Worksheet index in the package.
        /// -1             - Non-existing worksheet
        /// int.MinValue - Not set. 
        /// </summary>
        public int WorksheetIx;
        public int Row, Column;
        public string Address
        {
            get
            {
                if (Row > 0 && Column > 0)
                {
                    return ExcelAddressBase.GetAddress(Row, Column);
                }
                return "";
            }
        }
    }
    public class FormulaAddressBase
    {
        /// <summary>
        /// External reference index. 
        /// -1 means the current workbook.
        /// short.MinValue - Invalid reference.
        /// </summary>
        public int ExternalReferenceIx = -1;
        /// <summary>
        /// Worksheet index in the package.
        /// -1             - Non-existing worksheet
        /// short.MinValue - Not set. 
        /// </summary>
        public int WorksheetIx = int.MinValue;
    }
    public class FormulaRangeAddress : FormulaAddressBase, IAddressInfo, IComparable<FormulaRangeAddress>
    {
        public ParsingContext _context;
        internal FormulaRangeAddress()
        {

        }
        internal FormulaRangeAddress(ParsingContext ctx)
        {            
            _context = ctx;
            if(WorksheetIx==int.MinValue && ctx!=null) 
            {
                WorksheetIx = ctx.CurrentCell.WorksheetIx;
            }
        }
        internal FormulaRangeAddress(ParsingContext ctx, string address) : this(ctx)
        {
            int ix;
            if (address.StartsWith("["))
            {
                ix = address.IndexOf(']');
                if(ix>1)
                {
                    ExternalReferenceIx = (short)ctx.Package.Workbook.ExternalLinks.GetExternalLink(address.Substring(1, ix - 1));
                    address = address.Substring(ix + 1);
                }
            }
            ix = address.LastIndexOf('!');
            while (ix > 0)
            {
                if ((ix > 4 && address.Substring(ix - 4, 4).Equals("#REF!")) == false) break;
                address.LastIndexOf('!', ix-1);
            }
            if(ix>0)
            {
                var ws = address.Substring(0, ix - 1);
                address = address.Substring(ix + 1);
                if(ws.StartsWith("'") && ws.EndsWith("'"))
                {
                    ws = ws.Substring(1, ws.Length - 2).Replace("''", "'");
                }
            }
            ExcelCellBase.GetRowColFromAddress(address, out FromRow, out FromCol, out ToRow, out ToCol);
        }
        public int FromRow, FromCol, ToRow, ToCol;
        internal FixedFlag FixedFlag;

        public bool IsSingleCell
        {
            get
            {
                return FromRow == ToRow && FromCol == ToCol;
            }
        }
        public static FormulaRangeAddress Empty
        {
            get { return new FormulaRangeAddress(); }
        }

        internal eAddressCollition CollidesWith(FormulaRangeAddress other)
        {
            var util = new ExcelAddressCollideUtility(this, _context);
            return util.Collide(other, _context);
        }

        /// <summary>
        /// ToString() returns the full address as a string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var ws = WorksheetName;
            if(!string.IsNullOrEmpty(ws))
            {
                return new ExcelAddress(ws, FromRow, FromCol, ToRow, ToCol).FullAddress;
            }
            return new ExcelAddress(FromRow, FromCol, ToRow, ToCol).FullAddress;
        }

        /// <summary>
        /// Address of the range on the worksheet (i.e. worksheet name is excluded).
        /// </summary>
        public string WorksheetAddress
        {
            get
            {
                return new ExcelAddress(FromRow, FromCol, ToRow, ToCol).Address;
            }
        }

        /// <summary>
        /// Worksheet name of the address
        /// </summary>
        public string WorksheetName
        {
            get
            {
                if(WorksheetIx > -1 && _context != null && _context.Package != null)
                {
                    if(_context.Package.Workbook.Worksheets[WorksheetIx] != null)
                    {
                        return _context.Package.Workbook.Worksheets[WorksheetIx].Name;
                    }
                }
                return string.Empty;
            }
        }

        internal FormulaRangeAddress Intersect(FormulaRangeAddress address)
        {
            if (address.FromRow > ToRow || ToRow < address.FromRow ||
               address.FromCol > ToCol || ToCol < address.FromCol ||
               address.WorksheetIx != WorksheetIx)
            {
                return null;
            }

            var fromRow = Math.Max(address.FromRow, FromRow);
            var toRow = Math.Min(address.ToRow, ToRow);
            var fromCol = Math.Max(address.FromCol, FromCol);
            var toCol = Math.Min(address.ToCol, ToCol);

            return new FormulaRangeAddress(_context)
            {
                WorksheetIx = WorksheetIx,
                FromRow = fromRow,
                FromCol = fromCol,
                ToRow = toRow,
                ToCol = toCol
            };
        }

        /// <summary>
        /// Returns this address as a <see cref="ExcelAddressBase"/>
        /// </summary>
        /// <returns></returns>
        internal ExcelAddressBase ToExcelAddressBase()
        {
            if(ExternalReferenceIx > 0)
            {
                return new ExcelAddressBase(ExternalReferenceIx, WorksheetName, FromRow, FromCol, ToRow, ToCol);
            }
            return new ExcelAddressBase(WorksheetName, FromRow, FromCol, ToRow, ToCol);
        }

        public int CompareTo(FormulaRangeAddress other)
        {
            if(FromRow < other.FromRow)
            {
                return -1;
            }
            else if(FromRow> other.FromRow)
            {
                return 1;
            }
            else
            {
                if(FromCol < other.FromCol)
                {
                    return -1;
                }
                else if (FromCol>other.FromCol)
                {
                    return 1;
                }
                return 0;
            }
        }
        public virtual FormulaRangeAddress Clone()
        {
            return new FormulaRangeAddress(_context)
            { 
                 ExternalReferenceIx= ExternalReferenceIx,
                 WorksheetIx= WorksheetIx,
                 FixedFlag= FixedFlag,
                 FromRow= FromRow,
                 FromCol= FromCol,
                 ToRow= ToRow,
                 ToCol= ToCol             
            };
        }
        internal bool CollidesWith(int wsIx, int row, int column)
        {
            return wsIx==WorksheetIx && row >= FromRow && row <= ToRow && column >= FromCol && column <= ToCol;
        }
        public FormulaRangeAddress Address => this;
    }
    public class FormulaTableAddress : FormulaRangeAddress
    {
        public FormulaTableAddress(ParsingContext ctx) 
        {
            _context = ctx;
        }
        public FormulaTableAddress(ParsingContext ctx, string tableAddress)
        {
            foreach (var t in OptimizedSourceCodeTokenizer.Default.Tokenize(tableAddress))
            {
                switch (t.TokenType)
                {
                    case TokenType.TableName:
                        TableName = t.Value;
                        break;
                    case TokenType.TableColumn:
                        if (string.IsNullOrEmpty(ColumnName1))
                        {
                            ColumnName1 = t.Value;
                        }
                        else
                        {
                            ColumnName2 = t.Value;
                        }
                        break;

                    case TokenType.TablePart:
                        if(string.IsNullOrEmpty(TablePart1))
                        {
                            TablePart1 = t.Value;
                        }   
                        else
                        {
                            TablePart2 = t.Value;
                        }
                        break;
                }
            }
            SetTableAddress(ctx.Package);
        }
        public string TableName = "", ColumnName1 = "", ColumnName2 = "", TablePart1 = "", TablePart2="";
        internal void SetTableAddress(ExcelPackage package)
        {
            ExcelTable table;
            if (WorksheetIx >= 0)
            {
                if(WorksheetIx< package.Workbook.Worksheets.Count)
                {
                    table = package.Workbook.Worksheets[WorksheetIx].Tables[TableName];
                }
                else
                {
                    table = null;
                }
            }
            else if(WorksheetIx == int.MinValue)
            {
                table = package.Workbook.GetTable(TableName);
                WorksheetIx = table.WorkSheet.IndexInList;
            }
            else
            {
                table = null;
            }

            if (table != null && ExternalReferenceIx <= 0)
            {
                FixedFlag = FixedFlag.All;

                FromRow = table.ShowHeader ? table.Address._fromRow + 1 : table.Address._fromRow;
                ToRow = table.ShowTotal ? table.Address._toRow - 1 : table.Address._toRow;
                FromCol = table.Address._fromCol;
                ToCol = table.Address._toCol;

                SetRowFromTablePart(TablePart1, table, ref FromRow, ref ToRow, ref FixedFlag);
                if(string.IsNullOrEmpty(TablePart2)==false) SetRowFromTablePart(TablePart2, table, ref FromRow, ref ToRow, ref FixedFlag);
                
                SetColFromTablePart(ColumnName1, table, ref FromCol, ref ToCol, false);
                if (string.IsNullOrEmpty(ColumnName2) == false) SetColFromTablePart(ColumnName2, table, ref FromCol, ref ToCol, true);
            }
        }
        private void SetColFromTablePart(string value, ExcelTable table, ref int fromCol, ref int toCol, bool lastColon)
        {            
            var col = table.Columns[value];
            if (col == null)
            {
                if(value.StartsWith("'#"))
                {
                    col = table.Columns[value.Substring(1)];
                }
                if (col == null)
                    return;
            }
            if (lastColon)
            {
                toCol = table.Range._fromCol + col.Position;
            }
            else
            {
                fromCol = toCol = table.Range._fromCol + col.Position;
            }
        }
        private void SetRowFromTablePart(string value, ExcelTable table, ref int fromRow, ref int toRow, ref FixedFlag fixedFlag)
        {
            switch (value.ToLower())
            {
                case "#all":
                    fromRow = table.Address._fromRow;
                    toRow = table.Address._toRow;
                    break;
                case "#headers":
                    if (table.ShowHeader)
                    {
                        fromRow = table.Address._fromRow;
                        if (toRow == 0)
                        {
                            toRow = table.Address._fromRow;
                        }
                    }
                    else if (fromRow == 0)
                    {
                        fromRow = toRow = -1;
                    }
                    break;
                case "#data":
                    if (fromRow == 0 || table.DataRange._fromRow < fromRow)
                    {
                        fromRow = table.DataRange._fromRow;
                    }
                    if (table.DataRange._toRow > toRow)
                    {
                        toRow = table.DataRange._toRow;
                    }
                    break;
                case "#totals":
                    if (table.ShowTotal)
                    {
                        if (fromRow == 0)
                            fromRow = table.Range._toRow;
                        toRow = table.Range._toRow;
                    }
                    else if (fromRow == 0)
                    {
                        fromRow = toRow = -1;
                    }
                    break;
                case "#this row":
                    var dr = table.DataRange;
                    var r = _context.CurrentCell.Row;
                    if (WorksheetIx != table.WorkSheet.IndexInList || r < dr._fromRow || r > dr._toRow)
                    {
                        fromRow = toRow = -1;
                    }
                    else
                    {
                        fromRow = r;
                        toRow = r;
                        fixedFlag = FixedFlag.FromColFixed | FixedFlag.ToColFixed;
                    }
                    break;
                default:
                    FromCol = table.Address._fromCol;
                    ToCol = table.Address._toCol;
                    break;
            }
        }
        public virtual FormulaTableAddress Clone()
        {
            return new FormulaTableAddress(_context)
            {
                ExternalReferenceIx = ExternalReferenceIx,
                WorksheetIx = WorksheetIx,
                FixedFlag = FixedFlag,
                FromRow = FromRow,
                FromCol = FromCol,
                ToRow = ToRow,
                ToCol = ToCol,
                TableName = TableName,
                TablePart1 = TablePart1,
                TablePart2 = TablePart2,
                ColumnName1 = ColumnName1,
                ColumnName2 = ColumnName2,
            };
        }
    }
}
