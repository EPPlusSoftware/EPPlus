using OfficeOpenXml.DataValidation.Exceptions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/07/2024         EPPlus Software AB       EPPlus 7.2
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable
{
	internal class PivotTableColumnCalculation
	{
		ExcelPivotTable _tbl;
		List<PivotCalculationStore> _calcItems;
		List<int> _calcOrder;
		FormulaParser _formulaParser;
		FunctionRepository _fr;
        public PivotTableColumnCalculation(ExcelPivotTable tbl)
        {
			_tbl = tbl;
			_formulaParser = _tbl.WorkSheet.Workbook.FormulaParser;
			_fr = _tbl.WorkSheet.Workbook.FormulaParser.ParsingContext.Configuration.FunctionRepository;
			_calcItems = tbl.CalculatedItems;
        }
		internal void CalculateFormulaFields(List<int> fieldIndex)
		{
			var calcOrder = GetCalcOrder();
			foreach(var i in calcOrder)
			{
				var f = _tbl.Fields[i];
				var tokens = (List<Token>)SourceCodeTokenizer.PivotFormula.Tokenize(f.CacheField.Formula);
				var calcTokens = GetPivotFieldReferencesInFormula(f, tokens);
				if(calcTokens.Any(x => x == null))
				{
					//Contains invalid field reference or functions not supported in PT.

				}
				else
				{
					CalculateField(f, tokens, calcTokens, fieldIndex);
				}
			}
		}

		private void CalculateField(ExcelPivotTableField f, List<Token> tokens, List<int[]> calcTokens, List<int> fieldIndex)
		{
			var options = new ExcelCalculationOption();
			var depChain = new RpnOptimizedDependencyChain(_tbl.WorkSheet.Workbook, options);
			foreach (var ci in _tbl.CalculatedItems[0].Index)
			{
				foreach (var i in calcTokens)
				{
					var v = _tbl.CalculatedItems[fieldIndex[i[1]]][ci.Key];
					tokens[i[0]] = GetTokenFromValue(v);
				}
				RpnFormulaExecution.ExecutePivotFieldFormula(depChain, tokens, options);
			}
		}

		private Token GetTokenFromValue(object v)
		{
			if(ConvertUtil.IsNumericOrDate(v))
			{
				return new Token(ConvertUtil.GetValueDouble(v).ToString(CultureInfo.InvariantCulture), TokenType.Decimal);
			}
			else if(v is ExcelErrorValue ev)
			{
				switch(ev.Type)
				{
					case eErrorType.Ref:
						return new Token(ev.ToString(), TokenType.InvalidReference);
					case eErrorType.NA:
						return new Token(ev.ToString(), TokenType.NAError);
					case eErrorType.Num:
						return new Token(ev.ToString(), TokenType.NumericError);
					default:
						return new Token(ev.ToString(), TokenType.ValueDataTypeError);
				}
			}
			return new Token(v.ToString(),TokenType.String);
		}

		private List<int[]> GetPivotFieldReferencesInFormula(ExcelPivotTableField f, IList<Token> tokens)
		{
			var ret = new List<int[]>();
			int ix = 0;
			foreach (var t in tokens)
			{
				if(t.TokenType==TokenType.NameValue)
				{
					var ff = _tbl.Fields[t.Value];
					if (ff == null)
					{
						ret.Add(null);
						return ret;
					}
					else
					{
						ret.Add([ix, ff.Index]);
					}
				}
				else if(
					t.TokenType == TokenType.Array ||
				    t.TokenType == TokenType.CellAddress ||
				    t.TokenType == TokenType.FullColumnAddress ||
					t.TokenType == TokenType.FullRowAddress ||
					t.TokenType == TokenType.TableName ||
					t.TokenType == TokenType.WorksheetName)
				{
					ret.Add(null);
					return ret;
				}
				else if(t.TokenType==TokenType.Function)
				{
					var function = _fr.GetFunction(t.Value);
					if(function.IsAllowedInCalculatedPivotTableField==false)
					{
						ret.Add(null);
						return ret;
					}
				}
				ix++;
			}
			return ret;
		}

		private List<int> GetCalcOrder()
		{
			var calcOrder = new List<int>();
			foreach (var f in _tbl.Fields.Where(x => string.IsNullOrEmpty(x.CacheField.Formula) == false))
			{
				if (calcOrder.Contains(f.Index)) continue;
				ValidateNoCircularReference(f, calcOrder);
			}
			return calcOrder;
		}

		private bool ValidateNoCircularReference(ExcelPivotTableField f, List<int> calcOrder, Stack<ExcelPivotTableField> prevFields = null)
		{
			if (prevFields == null) prevFields = new Stack<ExcelPivotTableField>();
			var tokens = SourceCodeTokenizer.PivotFormula.Tokenize(f.CacheField.Formula);
			foreach (var t in tokens)
			{
				if (t.TokenType == TokenType.PivotField)
				{
					var f2 = _tbl.Fields[t.Value];
					if (f2 != null && string.IsNullOrEmpty(f2.Cache.Formula))
					{
						if (t.Value.Equals(f.Name, StringComparison.InvariantCultureIgnoreCase))
						{
							throw(new InvalidOperationException($"Circular reference in pivot table {_tbl.Name} Calculated Field {f.Name}"));
						}
						if(prevFields.Any(x=>x.Name.Equals(t.Value, StringComparison.InvariantCultureIgnoreCase)))
						{
							throw(new InvalidOperationException($"Circular reference in pivot table {_tbl.Name} Calculated Field {f.Name}"));
						}

						prevFields.Push(f);
						ValidateNoCircularReference(f2, calcOrder, prevFields);
					}
				}
			}
			calcOrder.Add(f.Index);
			return true;
		}
	}
}