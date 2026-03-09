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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
    internal class PivotTableColumnCalculation
	{
		ExcelPivotTable _tbl;
		List<PivotCalculationStore> _calcItems;
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
            var cacheFields = _tbl.CacheDefinition._cacheReference.Fields;

            foreach (var cfIndex in calcOrder)
            {
                var cf = cacheFields[cfIndex];
                var tokens = cf.FormulaTokens;

                // Build pivot field references from tokens
                // (mirrors original GetPivotFieldReferencesInFormula logic, but against cache)
                var calcTokens = new List<int[]>();
                bool hasInvalid = false;
                int ix = 0;
                foreach (var t in tokens)
                {
                    if (t.TokenType == TokenType.PivotField)
                    {
                        var refCf = cacheFields.FirstOrDefault(
                            x => x.Name.Equals(t.Value, StringComparison.InvariantCultureIgnoreCase));
                        if (refCf != null)
                        {
                            calcTokens.Add(new int[] { ix });
                        }
                        else
                        {
                            hasInvalid = true;
                            break;
                        }
                    }
                    else if (t.TokenType == TokenType.Array ||
                             t.TokenType == TokenType.CellAddress ||
                             t.TokenType == TokenType.FullColumnAddress ||
                             t.TokenType == TokenType.FullRowAddress ||
                             t.TokenType == TokenType.TableName ||
                             t.TokenType == TokenType.WorksheetName)
                    {
                        hasInvalid = true;
                        break;
                    }
                    else if (t.TokenType == TokenType.Function)
                    {
                        var func = _fr.GetFunction(t.Value);
                        if (func != null && func.IsAllowedInCalculatedPivotTableField == false)
                        {
                            hasInvalid = true;
                            break;
                        }
                    }
                    ix++;
                }

                PivotCalculationStore store;
                if (hasInvalid)
                {
                    throw new InvalidOperationException(
                        $"Pivot table {_tbl.Name} contains invalid column calculated formula : " +
                        $"{cf.Formula}. The formula contains an invalid field, an unsupported " +
                        "function or cell reference.");
                }
                else
                {
                    store = new PivotCalculationStore();
                    var options = new ExcelCalculationOption();
                    var depChain = new RpnOptimizedDependencyChain(_tbl.WorkSheet.Workbook, options);
                    var ct = new List<Token>();
                    ct.AddRange(tokens.Select(x => new Token(x.Value, x.TokenType, x.IsNegated)));

                    // Collect ALL unique keys across all referenced source fields,
                    // so we don't miss keys that only exist in some fields.
                    var allKeys = new List<int[]>();
                    var seenKeys = new HashSet<string>();
                    foreach (var c in calcTokens)
                    {
                        var fieldName = tokens[c[0]].Value;
                        if (_tbl.CalculatedFieldReferencedItems.TryGetValue(fieldName, out var refStore))
                        {
                            foreach (var k in refStore.Index)
                            {
                                var keyStr = string.Join(",", k.Key.Select(x => x.ToString()).ToArray());
                                if (seenKeys.Add(keyStr))
                                {
                                    allKeys.Add(k.Key);
                                }
                            }
                        }
                    }

                    // Calculate formula for each key combination
                    foreach (var key in allKeys)
                    {

                        foreach (var c in calcTokens)
                        {
                            var fieldName = tokens[c[0]].Value;
                            if (_tbl.CalculatedFieldReferencedItems.TryGetValue(fieldName, out var refStore)
                                && refStore.ContainsKey(key))
                            {
                                ct[c[0]] = GetTokenFromValue(refStore[key]);
                            }
                            else
                            {
                                // Key doesn't exist for this field — use 0 as default
                                // (Excel treats missing pivot values as 0 in calculated fields)
                                ct[c[0]] = new Token("0", TokenType.Decimal);
                            }
                        }
                        var cv = RpnFormulaExecution.ExecutePivotFieldFormula(depChain, ct, options);
                        store.Add(key, cv);
                    }
                }

                // Map back to DataField by matching the cache field.
                // Multiple DataFields can reference the same calculated cache field
                // (e.g. same field with different ShowDataAs), so don't break on first match.
                bool stored = false;
                for (int d = 0; d < _tbl.DataFields.Count; d++)
                {
                    var dfCache = _tbl.DataFields[d].Field.Cache;
                    if (dfCache == cf ||
                        dfCache.Name.Equals(cf.Name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        _tbl.CalculatedItems[d] = store;
                        stored = true;
                    }
                }
                if (!stored)
                {
                    _tbl.CalculatedFieldReferencedItems[cf.Name] = store;
                }
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
					case eErrorType.GettingData:
						return new Token(ev.ToString(), TokenType.GettingDataError);
                    case eErrorType.Num:
						return new Token(ev.ToString(), TokenType.NumericError);
					case eErrorType.Div0:
						return new Token(ev.ToString(), TokenType.Div0Error);
                    default:
						return new Token(ev.ToString(), TokenType.ValueDataTypeError);
				}
			}
			return new Token(v.ToString(),TokenType.String);
		}

        private List<int> GetCalcOrder()
        {
            var calcOrder = new List<int>();
            var cacheFields = _tbl.CacheDefinition._cacheReference.Fields;

            // Collect cache field indices that are used as DataFields and have formulas
            var relevantCfIndices = new HashSet<int>();
            foreach (var df in _tbl.DataFields)
            {
                if (!string.IsNullOrEmpty(df.Field.Cache.Formula))
                {
                    var cfIndex = cacheFields.IndexOf(df.Field.Cache);
                    if (cfIndex >= 0)
                    {
                        relevantCfIndices.Add(cfIndex);
                    }
                }
            }

            foreach (var cfIndex in relevantCfIndices)
            {
                if (calcOrder.Contains(cfIndex)) continue;
                ValidateNoCircularReferenceCacheField(
                    cacheFields[cfIndex], cfIndex, calcOrder, cacheFields);
            }
            return calcOrder;
        }

        private bool ValidateNoCircularReferenceCacheField(
            ExcelPivotTableCacheField cf,
            int cfIndex,
            List<int> calcOrder,
            List<ExcelPivotTableCacheField> cacheFields,
            Stack<int> prevIndices = null)
        {
            if (prevIndices == null) prevIndices = new Stack<int>();
            var tokens = SourceCodeTokenizer.PivotFormula.Tokenize(cf.Formula);
            foreach (var t in tokens)
            {
                if (t.TokenType == TokenType.PivotField)
                {
                    var refIndex = cacheFields.FindIndex(
                        x => x.Name.Equals(t.Value, StringComparison.InvariantCultureIgnoreCase));
                    if (refIndex >= 0 && !string.IsNullOrEmpty(cacheFields[refIndex].Formula))
                    {
                        if (refIndex == cfIndex || prevIndices.Contains(refIndex))
                        {
                            throw new InvalidOperationException(
                                $"Circular reference in pivot table {_tbl.Name} Calculated Field {cf.Name}");
                        }
                        prevIndices.Push(cfIndex);
                        ValidateNoCircularReferenceCacheField(
                            cacheFields[refIndex], refIndex, calcOrder, cacheFields, prevIndices);
                    }
                }
            }
            if (!calcOrder.Contains(cfIndex))
            {
                calcOrder.Add(cfIndex);
            }
            return true;
        }
	}
}