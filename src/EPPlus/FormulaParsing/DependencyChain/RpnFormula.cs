/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/14/2024         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.DependencyChain;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Security.AccessControl;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    internal enum RpnFormulaType
    {
        Formula,
        NameFormula,
        FixedArrayFormula
    }
    [Flags]
    internal enum FormulaFlags : short
    {
        IsDynamic           = 1,
        IsAllwaysDynamic    = 2,
    }
    internal class RpnFormula
    {
        internal ExcelWorksheet _ws;
        internal int _row;
        internal int _column;
        internal string _formula;
        internal RpnTokens _tokens;
        internal Dictionary<int, Expression> _expressions;
        internal int _enumeratorWorksheetIx;
        internal CellStoreEnumerator<object> _formulaEnumerator;
        internal int _tokenIndex = 0;
        internal Stack<Expression> _expressionStack;
        internal Stack<FunctionExpression> _funcStack;
        internal int _arrayIndex = -1;
        internal FormulaFlags _flags = 0;
        internal FunctionExpression _currentFunction = null;
        private VariableStorageManager _variableStorage;

        public bool CanBeDynamicArray
        {
            get
            {
                return _ws._flags.GetFlagValue(_row, _column, CellFlags.CanBeDynamicArray);
            }
        }

        internal VariableStorageManager VariableStorage => _variableStorage;


        internal RpnFormula(ExcelWorksheet ws, int row, int column)
            : this(ws, row, column, new VariableStorageManager())
        {
        }

        internal RpnFormula(ExcelWorksheet ws, int row, int column, VariableStorageManager variableStorage)
        {
            _ws = ws;
            _row = row;
            _column = column;
            _expressionStack = new Stack<Expression>();
            _funcStack = new Stack<FunctionExpression>();
            _variableStorage = variableStorage;
        }

        private LambdaFormulaSettings _lambdaSettings;
        internal LambdaFormulaSettings LambdaSettings
        {
            get
            {
                if(_lambdaSettings == null)
                {
                    _lambdaSettings = new LambdaFormulaSettings();
                }
                return _lambdaSettings;
            }
        }

        internal bool HasLambdaSettings => _lambdaSettings != null;

        internal bool HasLambdaToken(int tokenIx)
        {
            return _lambdaSettings != null && _lambdaSettings.LambdaTokens != null && _lambdaSettings.LambdaTokens.Contains(tokenIx);
        }

        internal LambdaExpressionStackPosition GetCurrentLambdaExpressionStackPosition()
        {
            if (_lambdaSettings == null || _lambdaSettings.CurrentLambdaExpressions == null) return null;
            return _lambdaSettings.CurrentLambdaExpressions.Count > 0 ? _lambdaSettings.CurrentLambdaExpressions.Peek() : null;
        }

        internal int GetNumberOfLambdaVariables()
        {
            if (_lambdaSettings == null || _lambdaSettings.NumberOfLambdaVariables == null || _lambdaSettings.NumberOfLambdaVariables.Count == 0) return 0;
            return _lambdaSettings.NumberOfLambdaVariables.Peek();
        }

        internal bool ShouldInvokeLambda(Stack<Expression> s)
        {
            var nLambdaArgs = GetNumberOfLambdaVariables();
            if(nLambdaArgs > 0)
            {
                return _lambdaSettings.CurrentLambdaExpressions.Count > 0 && (s.Count - _lambdaSettings.CurrentLambdaExpressions.Peek().StackIndex) == nLambdaArgs + 1;
            }
            return false;
        }

        internal void OnLambdaInvoked()
        {
            if(LambdaSettings != null)
            {
                if(LambdaSettings.CurrentLambdaExpressions.Count > 0)
                {
                    LambdaSettings.CurrentLambdaExpressions.Pop();
                }
                if(LambdaSettings.NumberOfLambdaVariables.Count > 0)
                {
                    LambdaSettings.NumberOfLambdaVariables.Pop();
                }
            }
        }

        internal short CurrentLambdaArgsAdded
        {
            get
            {
                if (_lambdaSettings == null) return 0;
                return _lambdaSettings.CurrentLambdaArgsAdded;
            }
        }

        internal string GetAddress()
        {

            if (_ws == null)
            {
                if (_row >= 0 && _column >= 0)
                {
                    return ExcelCellBase.GetAddress(_row, _column);
                }
                else
                {
                    return $"Workbook name - index {_row}";
                }
            }
            return _ws.Name + "!" + ExcelCellBase.GetAddress(_row, _column);
        }

        internal void SetFormula(string formula, RpnOptimizedDependencyChain depChain)
        {
            _tokens = FormulaExecutor.CreateRPNTokens(
                    depChain._tokenizer.Tokenize(formula));

            _formula = formula;
            _expressions = FormulaExecutor.CompileExpressions(this, ref _tokens, depChain._parsingContext);
        }
		internal void SetFormula(IList<Token> tokens, RpnOptimizedDependencyChain depChain)
		{
			_tokens = FormulaExecutor.CreateRPNTokens(tokens);
			_expressions = FormulaExecutor.CompileExpressions(ref _tokens, depChain._parsingContext);
		}

        internal void SetTokens(RpnTokens tokens, ParsingContext context)
        {
            _tokens = tokens;
            var formula = new StringBuilder();
            foreach (var token in tokens.Tokens)
            {
                formula.Append(token.Value);
            }
            _formula = formula.ToString();
            _expressions = FormulaExecutor.CompileExpressions(this, ref _tokens, context);
        }

        public override string ToString()
        {
            if (_ws == null)
            {
                return ExcelCellBase.GetAddress(_row, _column);
            }
            else
            {
                return _ws.Name + "!" + ExcelCellBase.GetAddress(_row, _column);
            }
        }

        internal void ClearCache()
        {
            foreach (var e in _expressions.Values)
            {
                if (e.ExpressionType == ExpressionType.CellAddress)
                    e._cachedCompileResult = null;
            }
        }

        internal virtual int GetWorksheetIndex()
        {
            return _ws.IndexInList;
        }
        internal virtual RpnFormulaType Type
        {
            get
            {
                return RpnFormulaType.Formula;
            }
        }
    }        
    internal class RpnNameFormula : RpnFormula
    {        
        internal RpnNameFormula(ExcelWorksheet ws, int row, int column, FormulaCellAddress currentCell) : base(ws, row, column)
        {
            CurrentCell = currentCell;
        }
        internal FormulaCellAddress CurrentCell { get;  }
        internal override int GetWorksheetIndex()
        {
            return CurrentCell.WorksheetIx;
        }
        internal override RpnFormulaType Type
        {
            get
            {
                return RpnFormulaType.NameFormula;
            }
        }
    }
    internal class RpnArrayFormula : RpnFormula
    {
        internal RpnArrayFormula(ExcelWorksheet ws, int startRow, int startColumn, int endRow, int endCol) : base(ws, startRow, startColumn)
        {
            _endRow = endRow;
            _endCol = endCol;
        }
        internal int _endRow, _endCol;
        internal override RpnFormulaType Type
        {
            get
            {
                return RpnFormulaType.FixedArrayFormula;
            }
        }
    }
}