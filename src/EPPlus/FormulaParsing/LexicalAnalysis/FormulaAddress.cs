using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;
using static OfficeOpenXml.ExcelAddressBase;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Linq;
using OfficeOpenXml.Utils;
using System.Net;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    internal class Formula
    {
        internal ExcelWorksheet _ws;
        internal int StartRow, StartCol;
        internal int StartRowOffset, StartColOffset; //If the shared formula does not begin on the top-left cell, this contains the offset to the row/column to the top left cell.
        internal static ISourceCodeTokenizer _tokenizer = SourceCodeTokenizer.Default;
        internal static ISourceCodeTokenizer _tokenizerNWS = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false, true);
        internal IList<Token> Tokens;
        internal IList<Token> RpnTokens;
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
            RpnTokens = FormulaExecutor.CreateRPNTokens(Tokens);
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
        public SharedFormula() : base()
        {
        }
        public SharedFormula(ExcelWorksheet ws) : base(ws)
        {
        }
        public SharedFormula(ExcelWorksheet ws, int row, int col, string address, string formula) : base(ws)
        {
            _ws = ws;
            Formula = formula;
            ExcelCellBase.GetRowColFromAddress(address, out int sr, out int sc, out EndRow, out EndCol); //We don't use the start row/col from the address as it can differ from the cells row/col if the first cell has been deleted. 
            StartRow = row;
            StartCol = col;
            StartRowOffset = sr - StartRow;
            StartColOffset = sc - StartCol;
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
                _hasUpdatedNamespace = false;
                SetFormula(_ws, value);
            }
        }
        internal bool IsArray { get; set; }
        internal string Address
        {
            get
            {
                return ExcelCellBase.GetAddress(StartRow + StartRowOffset, StartCol + StartColOffset, EndRow, EndCol);
            }
            set
            {
                StartRowOffset = 0;
                StartColOffset = 0;
                ExcelCellBase.GetRowColFromAddress(value, out StartRow, out StartCol, out EndRow, out EndCol); 
            }
        }
        /// <summary>
        /// Return all addresses in the formula.
        /// </summary>
        public List<string> TokenAddresses 
        {
            get
            {
                var l = new List<string>();
                var address = "";
                foreach(var t in Tokens)
                {
                    if(t.TokenTypeIsAddress || (t.Value==":" && t.TokenType==TokenType.Operator))
                    {
                        address += t.Value;
                    }
                    else
                    {
                        if(!string.IsNullOrEmpty(address))
                        {
                            l.Add(address);
                            address = "";
                        }
                    }
                }
                if (!string.IsNullOrEmpty(address))
                {
                    l.Add(address);
                }
                return l;
            }
        }
        /// <summary>
        /// Return tokens with addresses concatenated into an ExcelAddress instead of cell
        /// </summary>
        public List<Token> TokensWithFullAddresses
        {
            get
            {
                var l = new List<Token>();
                var address = "";
                foreach (var t in Tokens)
                {
                    if (t.TokenTypeIsAddress || (t.Value == ":" && t.TokenType == TokenType.Operator))
                    {
                        address += t.Value;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(address))
                        {
                            l.Add(new Token(address, TokenType.ExcelAddress));
                            address = "";                            
                        }
                        l.Add(t);
                    }
                }
                if (!string.IsNullOrEmpty(address))
                {
                    l.Add(new Token(address, TokenType.ExcelAddress));
                }
                return l;
            }
        }

        internal SharedFormula Clone()
        {
            return new SharedFormula(_ws, StartRow, StartCol, EndRow, EndCol, Formula)
            {
                Index = Index,
                FormulaType = FormulaType,
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
                //_compiler = _compiler,
            };
        }
        //RpnCompiledFormula _compiledExpressions = null;
        private Dictionary<int, Expression> _compiledExpressions;
        internal RpnFormula GetRpnFormula(RpnOptimizedDependencyChain depChain, int row, int col)
        {
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(_ws.IndexInList, row, col);
            if (_compiledExpressions == null)
            {
                _compiledExpressions = FormulaExecutor.CompileExpressions(ref RpnTokens, depChain._parsingContext);
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
        internal RpnFormula GetRpnArrayFormula(RpnOptimizedDependencyChain depChain, int startRow, int startCol, int endRow, int endCol)
        {
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(_ws.IndexInList, startRow, startCol);
            if (_compiledExpressions == null)
            {
                _compiledExpressions = FormulaExecutor.CompileExpressions(ref RpnTokens, depChain._parsingContext);
            }
            return new RpnArrayFormula(_ws, startRow, startCol, endRow, endCol)
            {
                _tokenIndex = 0,
                _tokens = RpnTokens,
                _expressions = CloneExpressions(startRow, startCol)
            };
        }

        private Dictionary<int, Expression> CloneExpressions(int row, int col)
        {
            var l=new Dictionary<int, Expression>();
            foreach(var expression in _compiledExpressions)
            {
                if(expression.Value.ExpressionType == ExpressionType.CellAddress ||
                   expression.Value.ExpressionType == ExpressionType.TableAddress ||
                   expression.Value.ExpressionType == ExpressionType.NameValue)
                {
                    l.Add(expression.Key, expression.Value.CloneWithOffset(row - StartRow, col - StartCol));
                }
                else if(expression.Value.ExpressionType == ExpressionType.Function)
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
                    f += token.Value;
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

        internal bool _hasUpdatedNamespace=false;
        internal void UpdateFormulaNamespaces(Dictionary<string, string> nsDict)
        {
            _formula = UpdateFormulaNamespaces(_formula, nsDict);
            _hasUpdatedNamespace = true;
        }

        internal static string UpdateFormulaNamespaces(string formula, Dictionary<string, string> nsDict)
        {
            if (nsDict.Keys.Any(x => formula.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                var sb = new StringBuilder();
                var tokens = _tokenizerNWS.Tokenize(formula);
                foreach (var t in tokens)
                {
                    if (t.TokenTypeIsSet(TokenType.Function) && nsDict.ContainsKey(t.Value))
                    {
                        sb.Append(nsDict[t.Value] + t.Value);
                    }
                    else
                    {
                        sb.Append(t.Value);
                    }
                }
                return sb.ToString();
            }
            return formula;
        }
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
    /// <summary>
    /// Formula Cell address
    /// </summary>
    public struct FormulaCellAddress
    {
        /// <summary>
        /// Constructor cell address
        /// </summary>
        /// <param name="wsIx"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
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
        /// <summary>
        /// The row number
        /// </summary>
        public int Row;
        /// <summary>
        /// The column number
        /// </summary>
        public int Column;
        /// <summary>
        /// The address
        /// </summary>
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

        /// <summary>
        /// The cell id for the address. 
        /// The cell Id is an ulong with the worksheet shifted as <code>((ushort)sheetId) | (((ulong)col) &lt;&lt; 16) | (((ulong)row) &lt;&lt; 30)</code>
        /// </summary>
        public ulong CellId 
        { 
            get 
            { 
                return ExcelCellBase.GetCellId(WorksheetIx,Row,Column);
            } 
        }
    }
    /// <summary>
    /// Formula address base
    /// </summary>
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
    internal struct SimpleAddress
    {
        internal SimpleAddress(int fromRow, int fromCol, int toRow, int toCol)
        {
            FromRow = fromRow;
            FromCol = fromCol;
            ToRow = toRow;
            ToCol = toCol;
        }
        internal int FromRow, FromCol, ToRow,ToCol;
    }
    /// <summary>
    /// Represents a range address
    /// </summary>
    public class FormulaRangeAddress : FormulaAddressBase, IAddressInfo, IComparable<FormulaRangeAddress>
    {
        internal ParsingContext _context;
        /// <summary>
        /// Constructor
        /// </summary>
        public FormulaRangeAddress()
        {

        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ctx"></param>
        public FormulaRangeAddress(ParsingContext ctx)
        {            
            _context = ctx;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="address"></param>
        public FormulaRangeAddress(ParsingContext ctx, string address) : this(ctx)
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
                var ws = address.Substring(0, ix);
                address = address.Substring(ix + 1);
                if(ws.StartsWith("'") && ws.EndsWith("'"))
                {
                    ws = ws.Substring(1, ws.Length - 2).Replace("''", "'");
                }
                WorksheetIx=ctx.GetWorksheetIndex(ws);
            }
            else if (WorksheetIx == int.MinValue && ctx != null)
            {                
                WorksheetIx = ctx.CurrentCell.WorksheetIx;
            }
            ExcelCellBase.GetRowColFromAddress(address, out FromRow, out FromCol, out ToRow, out ToCol, 
                out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);

            FixedFlag = fixedFromRow ? FixedFlag.FromRowFixed : 0;
            FixedFlag |= fixedFromCol ? FixedFlag.FromColFixed : 0;
            FixedFlag |= fixedToRow ? FixedFlag.ToRowFixed : 0;
            FixedFlag |= fixedToCol ? FixedFlag.ToColFixed : 0;
        }
        /// <summary>
        /// Formula range address
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="address"></param>
        public FormulaRangeAddress(ParsingContext ctx, ExcelAddressBase address) : this(ctx)
        {
            FromRow = address._fromRow;
            FromCol = address._fromCol;
            ToRow = address._toRow;
            ToCol = address._toCol;

            FixedFlag = address._fromRowFixed ? FixedFlag.FromRowFixed : 0;
            FixedFlag |= address._fromColFixed ? FixedFlag.FromColFixed : 0;
            FixedFlag |= address._toRowFixed ? FixedFlag.ToRowFixed : 0;
            FixedFlag |= address._toColFixed ? FixedFlag.ToColFixed : 0;
            if(ctx!=null)
            {
                WorksheetIx = ctx.GetWorksheetIndex(address.WorkSheetName);
            }
        }
        /// <summary>
        /// Formula range address
        /// </summary>
        /// <param name="context"></param>
        /// <param name="wsIx"></param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="toRow"></param>
        /// <param name="toCol"></param>
        public FormulaRangeAddress(ParsingContext context,int wsIx, int fromRow, int fromCol, int toRow, int toCol) : this(context)
        {
            WorksheetIx= wsIx;
            FromRow = fromRow;
            FromCol = fromCol;
            ToRow = toRow;
            ToCol = toCol;
        }
        /// <summary>
        /// From row and column. To row and to column
        /// </summary>
        public int FromRow, FromCol, ToRow, ToCol;
        internal FixedFlag FixedFlag;
        /// <summary>
        /// Is single cell
        /// </summary>
        public bool IsSingleCell
        {
            get
            {
                return FromRow == ToRow && FromCol == ToCol;
            }
        }
        /// <summary>
        /// Empty
        /// </summary>
        public static FormulaRangeAddress Empty
        {
            get { return new FormulaRangeAddress(); }
        }

        internal eAddressCollition CollidesWith(FormulaRangeAddress other)
        {
            var util = new ExcelAddressCollideUtility(this, _context);
            return util.Collide(other, _context);
        }
        internal bool DoCollide(List<SimpleAddress> addresses)
        {            
            foreach(var a in addresses)
            {
                if(DoCollide(a.FromRow, a.FromCol, a.ToRow, a.ToCol))
                {
                    return true;
                }
            }
            return false;
        }
        internal bool DoCollide(int fromRow, int fromCol, int toRow, int toCol)
        {
            return fromRow <= ToRow && fromCol <= ToCol
                   &&
                   FromRow <= toRow && FromCol <= toCol;
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
                return ExcelCellBase.GetAddress(FromRow, FromCol, ToRow, ToCol, EnumUtil.HasFlag(FixedFlag, FixedFlag.FromRowFixed), EnumUtil.HasFlag(FixedFlag, FixedFlag.FromColFixed), EnumUtil.HasFlag(FixedFlag, FixedFlag.ToRowFixed), EnumUtil.HasFlag(FixedFlag, FixedFlag.ToColFixed));
            }
        }

        /// <summary>
        /// Worksheet name of the address
        /// </summary>
        public string WorksheetName
        {
            get
            {
                if(WorksheetIx > -1 && ExternalReferenceIx > 0)
                {
                    if (_context.Package.Workbook.ExternalLinks.Count >= ExternalReferenceIx)
                    {
                        var ewb = _context.Package.Workbook.ExternalLinks[ExternalReferenceIx - 1].As.ExternalWorkbook;
                        if (ewb.CachedWorksheets.Count > WorksheetIx)
                        {
                            return ewb.CachedWorksheets[WorksheetIx].Name;
                        }
                    }
                }
                else if(WorksheetIx > -1 && _context != null && _context.Package != null)
                {
                    if(_context.Package.Workbook.Worksheets.Count > WorksheetIx)
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
        /// <summary>
        /// Compare to
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
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
        /// <summary>
        /// Clone
        /// </summary>
        /// <returns></returns>
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

        internal FormulaRangeAddress GetOffset(int row, int column, bool rollIfOverflow=false)
        {
            var ret = new FormulaRangeAddress(_context);
            ret.ExternalReferenceIx = ExternalReferenceIx;
            ret.WorksheetIx = WorksheetIx;
            var fromRow = (FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.FromRowFixed ? FromRow : FromRow + row-1;            
            var toRow = (FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.ToRowFixed ? ToRow : ToRow + row-1;
            var fromCol = (FixedFlag & FixedFlag.FromColFixed) == FixedFlag.FromColFixed ? FromCol : FromCol + column-1;
            var toCol = (FixedFlag & FixedFlag.ToColFixed) == FixedFlag.ToColFixed ? ToCol : ToCol + column - 1;            
            if(rollIfOverflow)
            {
                GetRolledValue(ref fromRow, ExcelPackage.MaxRows);
                GetRolledValue(ref toRow, ExcelPackage.MaxRows);
                GetRolledValue(ref fromCol, ExcelPackage.MaxColumns);
                GetRolledValue(ref toCol, ExcelPackage.MaxColumns);
            }
            ret.FromRow = fromRow;
            ret.ToRow = toRow;
            ret.FromCol = fromCol;
            ret.ToCol = toCol;
            ret.FixedFlag = FixedFlag;
            return ret;
        }
        internal FormulaRangeAddress GetOffset(int row, int column, int rows, int columns)
        {
            var ret = new FormulaRangeAddress(_context);
            ret.ExternalReferenceIx = ExternalReferenceIx;
            ret.WorksheetIx = WorksheetIx;

            ret.FromRow = FromRow + row;
            ret.ToRow = ret.FromRow + rows - 1;
            ret.FromCol = FromCol + column;
            ret.ToCol = ret.FromCol + columns - 1;
            ret.FixedFlag = FixedFlag;
            return ret;
        }

        private void GetRolledValue(ref int value, int maxValue)
        {
            if(value < 1)
            {
                value = Math.Abs(value) + 1;
            }
            else if(value > maxValue)
            {
                value = value - maxValue;
            }            
        }

        internal ulong GetTopLeftCellId()
        {
            return ExcelCellBase.GetCellId(WorksheetIx, FromRow, FromCol);
        }

        internal IRangeInfo GetAsRangeInfo()
        {
            if(ExternalReferenceIx>0)
            {
                return GetAsExternalRangeInfo();
            }
            else
            {
                return new RangeInfo(this);
            }
        }

        internal IRangeInfo GetAsExternalRangeInfo()
        {
            var wb = _context.GetExternalWoorkbook(ExternalReferenceIx);
            if(wb==null)
            {
                return new RangeInfo(null, -1, -1, -1, -1, _context, ExternalReferenceIx);
            }
            else if (wb.Package == null)
            {
                return new EpplusExcelExternalRangeInfo(wb, this, _context);
            }
            else
            {
                var ws = wb.Package.Workbook.GetWorksheetByIndexInList(WorksheetIx);
                var ri=new RangeInfo(ws, FromRow, FromCol, ToRow, ToCol, _context);
                ri.Address.ExternalReferenceIx = ExternalReferenceIx;
                return ri;
            }
        }
        /// <summary>
        /// Address
        /// </summary>
        public FormulaRangeAddress Address => this;
    }
    /// <summary>
    /// Formula table address
    /// </summary>
    public class FormulaTableAddress : FormulaRangeAddress
    {
        /// <summary>
        /// Formula table address constructor
        /// </summary>
        /// <param name="ctx"></param>
        public FormulaTableAddress(ParsingContext ctx) : base(ctx)
        {
            
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="tableAddress"></param>
        public FormulaTableAddress(ParsingContext ctx, string tableAddress) : base(ctx)
        {
            foreach (var t in SourceCodeTokenizer.Default.Tokenize(tableAddress))
            {
                switch (t.TokenType)
                {
                    case TokenType.TableName:
                        TableName = t.Value;
                        break;
                    case TokenType.TableColumn:
                        if (string.IsNullOrEmpty(ColumnName1))
                        {
                            ColumnName1 = ExcelTableColumn.DecodeTableColumnName(t.Value);
                        }
                        else
                        {
                            ColumnName2 = ExcelTableColumn.DecodeTableColumnName(t.Value);
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
        /// <summary>
        /// Names
        /// </summary>
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

            if (table != null)
            {
                FixedFlag = FixedFlag.All;

                if (string.IsNullOrEmpty(TablePart1) == false)
                {
                    SetRowFromTablePart(TablePart1, table, ref FromRow, ref ToRow, ref FixedFlag);
                    if (string.IsNullOrEmpty(TablePart2) == false)
                    {
                        SetRowFromTablePart(TablePart2, table, ref FromRow, ref ToRow, ref FixedFlag);
                    }
                }
                else
                {
                    FromRow = table.ShowHeader ? table.Address._fromRow + 1 : table.Address._fromRow;
                    ToRow = table.ShowTotal ? table.Address._toRow - 1 : table.Address._toRow;
                }

                if (string.IsNullOrEmpty(ColumnName1) == false)
                {
                    SetColFromTablePart(ColumnName1, table, ref FromCol, ref ToCol, false);
                    if (string.IsNullOrEmpty(ColumnName2) == false)
                    {
                        SetColFromTablePart(ColumnName2, table, ref FromCol, ref ToCol, true);
                    }
                }
                else
                {
                    FromCol = table.Address._fromCol;
                    ToCol = table.Address._toCol;
                }
            }
        }
        private void SetColFromTablePart(string value, ExcelTable table, ref int fromCol, ref int toCol, bool lastColon)
        {
            var col = table.Columns[value];
            if (col == null)
            {
                if(value.StartsWith("'#"))
                {
                    var colName = ConvertUtil.ExcelDecodeString(value.Substring(1));
                    col = table.Columns[colName];
                }
                if (col == null)
                {
                    fromCol = -1;
                    toCol = -1;
                    return;
                }
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
        /// <summary>
        /// Clones the table address.
        /// </summary>
        /// <returns></returns>
        public new FormulaTableAddress Clone()
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
