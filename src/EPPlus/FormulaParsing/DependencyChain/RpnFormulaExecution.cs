using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using static OfficeOpenXml.ExcelAddressBase;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.FormulaParsing
{
    internal class RpnFormulaExecution
    {
        internal static ArgumentParser _boolArgumentParser = new BoolArgumentParser();
        internal static bool _cacheExpressions = true;
        internal static RpnOptimizedDependencyChain Execute(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(wb, options);
            foreach (var ws in wb.Worksheets)
            {
                if (ws.IsChartSheet==false)
                {
                    ExecuteChain(depChain, ws.Cells, options, true);
                    ExecuteChain(depChain, ws.Names, options, true);
                }
            }
            ExecuteChain(depChain, wb.Names, options,  true);

            return depChain;
        }
        internal static RpnOptimizedDependencyChain Execute(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(ws.Workbook, options);
            ExecuteChain(depChain, ws.Cells, options, true);
            ExecuteChain(depChain, ws.Names, options, true);

            return depChain;
        }
        internal static RpnOptimizedDependencyChain Execute(ExcelRangeBase cells, ExcelCalculationOption options)
        {
            //Range chain
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(cells._workbook, options);

            if (cells is ExcelNamedRange name)
            {
                ExecuteName(depChain, name, options, true);
            }
            else
            {
                ExecuteChain(depChain, cells, options, true);
            }

            return depChain;
        }
        internal static object ExecuteFormula(ExcelWorksheet ws, string formula, ExcelCalculationOption options)
        {
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(ws.Workbook, options);
            return ExecuteChain(depChain, ws, formula, options, true);
        }
        internal static object ExecuteFormula(ExcelWorkbook wb, string formula, FormulaCellAddress cell, ExcelCalculationOption options)
        {
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(wb, options);
            ExcelWorksheet ws;
            if (cell.WorksheetIx < 0 || cell.WorksheetIx >= wb.Worksheets.Count)
            {
                ws = null;
            }
            else
            {
                ws = wb.Worksheets[cell.WorksheetIx];
            }
            return ExecuteChain(depChain, ws, formula, cell, options, false);
        }
        internal static object ExecuteFormula(ExcelWorkbook wb, string formula, ExcelCalculationOption options)
        {
            _cacheExpressions = options.CacheExpressions;
            var depChain = new RpnOptimizedDependencyChain(wb, options);

            return ExecuteChain(depChain, null, formula, options, true);
        }
		internal static object ExecutePivotFieldFormula(RpnOptimizedDependencyChain depChain, IList<Token> tokens, ExcelCalculationOption options)
		{
            var formula = new RpnFormula(null, 0, 0);
            formula.SetFormula(tokens, depChain);
			return AddChainForFormula(depChain, formula, options, false);
		}

		private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelRangeBase range, ExcelCalculationOption options, bool writeToCell)
        {
            var ws = range.Worksheet;
            RpnFormula f = null;
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
                {
                    try
                    {
                        if (GetFormula(depChain, ws, fs.Row, fs.Column, fs.Value, ref f))
                        {
                            AddChainForFormula(depChain, f, options, writeToCell);
                        }
                    }
                    catch (CircularReferenceException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        if(writeToCell)
                        {
                            SetAndReturnValueError(depChain, ex, f);
                        }
                    }
                }
            }
        }

        private static object SetAndReturnValueError(RpnOptimizedDependencyChain depChain, Exception ex, RpnFormula f)
        {
            if (depChain._parsingContext.Parser.Logger != null)
            {
                depChain._parsingContext.Parser.Logger.Log(depChain._parsingContext, ex);
                LogFormula(depChain, f);
            }
            var cc = depChain._parsingContext.CurrentCell;
            var ret = ExcelErrorValue.Create(eErrorType.Value);
            if (depChain._parsingContext.CurrentWorksheet!=null)
            {
                if(cc.Column>0)
                {
                    depChain._parsingContext.CurrentWorksheet.SetValueInner(cc.Row, cc.Column, ret);
                }
                else if (cc.Row >= 0 && cc.Row < depChain._parsingContext.CurrentWorksheet.Names.Count)
                {                    
                    depChain._parsingContext.CurrentWorksheet.Names[cc.Row].Value = ret;
                }
            }
            else if(cc.Column==0 && cc.Row >= 0 && cc.Row < depChain._parsingContext.Package.Workbook.Names.Count)
            {
                depChain._parsingContext.Package.Workbook.Names[depChain._parsingContext.CurrentCell.Row].Value = ret;
            }
            return ret;
        }

        private static void LogFormula(RpnOptimizedDependencyChain depChain, RpnFormula f)
        {
            try
            {
                var logger = depChain._parsingContext.Parser.Logger;
                logger.Log($"Formula at address: {f.GetAddress()}");
                logger.Log("Formula Tokens: " + string.Join(", ", f._tokens.Select(x => x.Value).ToArray()));
                logger.Log($"Formula current token : {f._tokens[f._tokenIndex]}. Position : {f._tokenIndex}");
                logger.Log($"Current Culture Setting: {Thread.CurrentThread.CurrentCulture.Name}");
            }
            catch (Exception)
            {

            }
        }

        private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelNamedRangeCollection namesCollection, ExcelCalculationOption options, bool writeToCell)
        {
            try 
            { 
                foreach (ExcelNamedRange name in namesCollection)
                {
                    ExecuteName(depChain, name, options, writeToCell);
                }
            }
            catch (CircularReferenceException)
            {
                throw;
            }
            catch (InvalidFormulaException ex)
            {
                depChain._parsingContext.Parser.Logger.Log(depChain._parsingContext, ex);
                throw;
            }
        }

        private static void ExecuteName(RpnOptimizedDependencyChain depChain, ExcelNamedRange name, ExcelCalculationOption options, bool writeToCell)
        {
            var ws = name._worksheet;
            var wsIx = ws == null ? -1 : ws.IndexInList;
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(wsIx, name.Index, 0);
            var id = ExcelCellBase.GetCellId(wsIx, name.Index, 0);
            if (depChain.processedCells.Contains(id) == false)
            {
                if (string.IsNullOrEmpty(name.NameFormula) == false)
                {
                    var f = GetNameFormula(depChain, ws, depChain._parsingContext.ExcelDataProvider.GetName(name),1,1);
                    AddChainForFormula(depChain, f, options, writeToCell);
                }
            }
        }

        private static object ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, string formula, FormulaCellAddress cell, ExcelCalculationOption options, bool writeToCell)
        {
            try 
            {
                var f = new RpnFormula(ws, cell.Row, cell.Column);
                f.SetFormula(formula, depChain);
                return AddChainForFormula(depChain, f, options, writeToCell);
            }
            catch (CircularReferenceException)
            {
                throw;
            }
            catch (InvalidFormulaException ex)
            {
                depChain._parsingContext.Parser.Logger.Log(depChain._parsingContext, ex);
                throw;
            }
        }

        private static object ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, string formula, ExcelCalculationOption options, bool writeToCell)
        {
            try 
            { 
                var f = new RpnFormula(ws, 0, 0);
                f.SetFormula(formula, depChain);
                f._row = -1;
                return AddChainForFormula(depChain, f, options, writeToCell);
            }
            catch (CircularReferenceException)
            {
                throw;
            }
            catch (InvalidFormulaException ex)
            {
                depChain._parsingContext.Parser.Logger.Log(depChain._parsingContext, ex);
                throw;
            }
        }
        private static bool GetFormula(RpnOptimizedDependencyChain depChain,  ExcelWorksheet ws, int row, int column, object value, ref RpnFormula f)
        {
            
            if (value is int ix)
            {
                var sf = ws._sharedFormulas[ix];
                if (sf.FormulaType==FormulaType.Array)
                {
                    MetaDataReference md = default;
                    bool isDynamic = false;
                    if (ws._metadataStore.Exists(sf.StartRow, sf.StartCol, ref md) && md.cm > 0)
                    {
                        isDynamic = ws.Workbook.Metadata.IsFormulaDynamic(md.cm);
                    }

                    if (isDynamic)
                    {
                        f = ws._sharedFormulas[ix].GetRpnFormula(depChain, sf.StartRow, sf.StartCol);
                        f._isDynamic = true;
                    }
                    else
                    {
                        f = ws._sharedFormulas[ix].GetRpnArrayFormula(depChain, sf.StartRow, sf.StartCol, sf.EndRow, sf.EndCol);
                    }
					f._arrayIndex = ix;
				}
				else
                {
                    f = ws._sharedFormulas[ix].GetRpnFormula(depChain, row, column);
                }
			}
			else
            {
                var s = value.ToString();
                //compiler
                if (string.IsNullOrEmpty(s)) return false;
                f = new RpnFormula(ws, row, column);
                SetCurrentCell(depChain, f);
                f.SetFormula(s, depChain);
            }
            f._ws._metadataStore.Clear(f._row, f._column, 1, 1);
            var id = ExcelCellBase.GetCellId(ws?.IndexInList ?? ushort.MaxValue, f._row, f._column);
            depChain.processedCells.Add(id);

            return true;
        }

        private static void SetCurrentCell(RpnOptimizedDependencyChain depChain, RpnFormula f)
        {
            if (f._column > 0)
            {
                depChain._parsingContext.CurrentCell = new FormulaCellAddress(f._ws.IndexInList, f._row, f._column);
            }
            else if(f.Type == RpnFormulaType.NameFormula)
            {
                var cc = ((RpnNameFormula)f).CurrentCell;
                if (cc.Row == 0) cc = new FormulaCellAddress(f._ws==null ? -1 : f._ws.IndexInList, f._row, f._column); //Not set, set to the name.
                depChain._parsingContext.CurrentCell =  cc;
            }
        }
        private static RpnFormula GetNameFormula(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, INameInfo name, int cellRow, int cellCol)
        {
            ExcelCellBase.SplitCellId(name.Id, out int wsIx, out int row, out int col);
            if (name.wsIx >= 0 && ws == null && depChain._parsingContext.Package.Workbook.Worksheets.Count > name.wsIx)
            {                
                ws = depChain._parsingContext.Package.Workbook.Worksheets[name.wsIx];
            }
            var f = new RpnNameFormula(ws, row , col, depChain._parsingContext.CurrentCell);
            if (cellRow == 0 || cellCol == 0)
            {
                f.SetFormula(name.Formula, depChain);
            }
            else
            {
                f.SetFormula(name.GetRelativeFormula(cellRow, cellCol), depChain);
            }
            return f;
        }
        private static object AddChainForFormula(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, bool writeToCell)
        {
            FormulaRangeAddress address = null;
            RangeHashset rd = AddAddressToRD(depChain, f._ws == null ? -1 : f._ws.IndexInList);
            object v=null;
            bool hasLogger = depChain._parsingContext.Parser.Logger != null;
            rd?.Merge(f._row, f._column);
            depChain.StartOfChain();
            var followChain = options.FollowDependencyChain;
        ExecuteFormula:
            try
            {
                SetCurrentCell(depChain, f);
                var ws = f._ws;

                if (f._tokenIndex < f._tokens.Count)
                {                    
                    address = ExecuteNextToken(depChain, f, followChain);
                    if (f._tokenIndex < f._tokens.Count)
                    {
                        if (address == null && f._expressions.ContainsKey(f._tokenIndex) && f._expressions[f._tokenIndex].ExpressionType == ExpressionType.NameValue)
                        {
                            var ne = f._expressions[f._tokenIndex] as NamedValueExpression;
                            if (ne._externalReferenceIx < 1)
                            {
                                rd = AddAddressToRD(depChain, ne._worksheetIx);

                                if (ne.IsRelative || rd.Merge(ExcelCellBase.GetRowFromCellId(ne._name.Id), 0))
                                {
                                    depChain._formulaStack.Push(f); 
                                    ws = ne._worksheetIx < 0 ? null : depChain._parsingContext.Package.Workbook._worksheets[ne._worksheetIx];
                                    
                                    f = GetNameFormula(depChain, ws, ((NamedValueExpression)f._expressions[f._tokenIndex])._name, f._row, f._column);
                                    goto ExecuteFormula;
                                }
                                else
                                {
                                    CheckCircularReferences(depChain, f, options);
                                    f._tokenIndex++;
                                    goto ExecuteFormula;
                                }
                            }
                            else
                            {
                                f._tokenIndex++;
                                goto ExecuteFormula;
                            }
                        }

                        if (address == null)
                        {
                            address = f._expressions[f._tokenIndex].GetAddress();
                        }
                        if (address.ExternalReferenceIx > 0) //We don't follow dep chain into external references.
                        {
                            f._tokenIndex++;
                            goto ExecuteFormula;
                        }
                        if (ws == null)
                        {
                            if (address?.WorksheetIx < 0)
                            {
                                throw new InvalidOperationException("Address in formula does not reference a worksheet and does not belong to a worksheet.");
                            }
                            else
                            {
                                ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                            }
                        }
                        else if (address?.WorksheetIx >= 0 && ws?.IndexInList != address?.WorksheetIx)
                        {
                            ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                        }

                        rd = AddAddressToRD(depChain, ws.IndexInList);

                        if (rd.Exists(address) || address.CollidesWith(ws.IndexInList, f._row, f._column))
                        {
                            CheckCircularReferences(depChain, f, address, options);
                        }

                        if (rd.ExistsGetSpill(ref address))
                        {
                            goto FollowChain;
                        }

                        f._tokenIndex++;
                        goto ExecuteFormula;
                    }
                }
                CompileResult cr;
                if (f._tokenIndex == int.MaxValue) //int.MaxValue means we have an invalid formulas and we should return a name error 
                {
                    cr = CompileResult.GetErrorResult(eErrorType.Name);
                }
                else
                {
                    cr = f._expressionStack.Pop().Compile();
                }

                if (writeToCell || depChain._formulaStack.Count > 0)  // If calculating single cell via the FormulaParser.Parse method we should not write to the cells
                {
                    SetValueToWorkbook(depChain, f, rd, cr);
                }

                if (hasLogger)
                {
                    depChain._parsingContext.Parser.Logger.Log($"Set value in Cell\t{f.GetAddress()}\t{cr.ResultValue}\t{cr.DataType}");
                }

                depChain.AddFormulaToChain(f);
                if (depChain._formulaStack.Count > 0)
                {
                    f = depChain._formulaStack.Pop();
                    if (f._formulaEnumerator == null)
                    {
                        f._tokenIndex++;
                        goto ExecuteFormula;
                    }
                    if (f._expressions.ContainsKey(f._tokenIndex))
                    {
                        address = f._expressions[f._tokenIndex].GetAddress();
                    }
                    else
                    {
                        address = f._expressionStack.Peek().GetAddress();
                    }                    
                    rd = AddAddressToRD(depChain, f._enumeratorWorksheetIx);
                    goto NextFormula;
                }
                return cr.ResultValue;
            FollowChain:
                ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                if (ws == null)
                {
                    f._tokenIndex++;
                    goto ExecuteFormula;
                }
                if (address.IsSingleCell)
                {
                    if (depChain.processedCells.Contains(ExcelCellBase.GetCellId(ws?.IndexInList??ushort.MaxValue, address.FromRow, address.FromCol)) == false)
                    {
                        rd?.Merge(address.FromRow, address.FromCol);
                        if (ws._formulas.Exists(address.FromRow, address.FromCol, ref v))
                        {
                            depChain._formulaStack.Push(f);
                            GetFormula(depChain, ws, address.FromRow, address.FromCol, v, ref f);
                            goto ExecuteFormula;
                        }
                    }
                    f._tokenIndex++;
                    goto ExecuteFormula;
                }
                else
                {
                    f._enumeratorWorksheetIx = ws.IndexInList;
                    f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
                }
            NextFormula:
                var fe = f._formulaEnumerator;
                var row = fe.Row;
                var col = fe.Column < 0 ? fe._startCol - 1 : fe.Column;
                if (fe.Next())
                {
                    if (fe.Value == null || depChain.processedCells.Contains(ExcelCellBase.GetCellId(f._enumeratorWorksheetIx, fe.Row, fe.Column)))
                    {
                        MergeToRd(rd, row, col, fe, false);
                        goto NextFormula;
                    }

                    depChain._formulaStack.Push(f);
                    MergeToRd(rd, row, col, fe, false);
                    if (GetFormula(depChain, ws, fe.Row, fe.Column, fe.Value, ref f))
                    {
                        goto ExecuteFormula;
                    }
                    else
                    {
                        goto NextFormula;
                    }
                }

                MergeToRd(rd, row, col, fe, true);

                f._formulaEnumerator = null;
                f._tokenIndex++;

                goto ExecuteFormula;
            }
            catch (CircularReferenceException)
            {
                throw;
            }
            catch (Exception ex)
            {
                object errValue;

                if (writeToCell)
                {
                    errValue = SetAndReturnValueError(depChain, ex, f);
                }
                else
                {
                    errValue = ExcelErrorValue.Create(eErrorType.Value);
                }

                f._tokenIndex=f._tokens.Count-1;
                if(depChain._formulaStack.Count > 0)
                {
                    f = depChain._formulaStack.Pop();
                    goto ExecuteFormula;
                }
                //goto CheckFormulaStack;
                return errValue;
            }

        }

        private static void SetValueToWorkbook(RpnOptimizedDependencyChain depChain, RpnFormula f, RangeHashset rd, CompileResult cr)
        {
            //Set the value.
            if (f._row >= 0)
            {
                if (f._ws == null)
                {                    
                    depChain._parsingContext.Package.Workbook.Names[f._row].SetValue(cr.ResultValue, depChain._parsingContext.CurrentCell);
                }
                else
                {
                    if (f._column == 0)
                    {
                        f._ws.Names[f._row].SetValue(cr.ResultValue, depChain._parsingContext.CurrentCell);
                    }
                    else
                    {
                        if ((cr.DataType == DataType.ExcelRange && ((IRangeInfo)cr.Result).Address.IsSingleCell==false)) //A range. When we add support for dynamic array formulas we will alter this.
                        {
                            var ri = (IRangeInfo)cr.Result;
                            if (f._arrayIndex >= 0 && f._isDynamic == false) //A legacy array formula, Fill the referenced range.
                            {
                                ArrayFormulaOutput.FillArrayFromRangeInfo(f, ri, rd, depChain);
                            }
                            else
                            {
                                if (f.CanBeDynamicArray) //Create a dynamic array formula if allowed. 
                                {
                                    //Add dynamic array formula support here.
                                    var dirtyRange = ArrayFormulaOutput.FillDynamicArrayFromRangeInfo(f, ri, rd, depChain);
                                    if (dirtyRange != null && dirtyRange.Length > 0)
                                    {

                                        RecalculateDirtyCells(dirtyRange, depChain, rd);
                                    }
                                }
                                else //Set implicit intersection
                                {
                                    var icr = ImplicitIntersectionUtil.GetResult(ri, f._row, f._column, depChain._parsingContext);
                                    f._ws.SetValueInner(f._row, f._column, icr.ResultValue ?? 0D);
                                }
                            }
                        }
                        else if (cr.ResultType == CompileResultType.DynamicArray)
                        {
                            var dirtyRange = ArrayFormulaOutput.FillDynamicArraySingleValue(f, cr, rd, depChain);
                            if (dirtyRange != null && dirtyRange.Length > 0)
                            {
                                RecalculateDirtyCells(dirtyRange, depChain, rd);
                            }
                        }
                        else
                        {
                            if(f._arrayIndex!=-1)
                            {
                                var sf = f._ws._sharedFormulas[f._arrayIndex];
                                f._ws.SetValueInner(sf.StartRow, sf.StartCol, sf.EndRow, sf.EndCol, cr.ResultValue ?? 0D);
                            }
                            else
                            {
                                f._ws.SetValueInner(f._row, f._column, cr.ResultValue ?? 0D);
                            }
                        }
                    }
                }
            }
        }

        private static void RecalculateDirtyCells(SimpleAddress[] dirtyRange, RpnOptimizedDependencyChain depChain, RangeHashset rd)
        {
            var dirtyCells = dirtyRange.ToList();
            foreach(var f in depChain._formulas)
            {
                foreach(var e in f._expressions.Values)
                {
                    if(e.Status==ExpressionStatus.IsAddress)
                    {
                        var a=e.GetAddress();                        
                        if(a.DoCollide(dirtyCells))
                        {
                            ReCalculateFormula(f, depChain, rd);
                            dirtyCells.Add(new SimpleAddress(a.FromRow, a.FromCol, a.ToRow, a.ToCol));
                        }
                    }
                }
            }
        }
        private static void ReCalculateFormula(RpnFormula f, RpnOptimizedDependencyChain depChain, RangeHashset rd)
        {
            f._tokenIndex = 0;
            f.ClearCache();
            ExecuteNextToken(depChain, f, false);
            var e=f._expressionStack.Pop();            
            SetValueToWorkbook(depChain, f, rd, e.Compile());
        }
        private static void MergeToRd(RangeHashset rd, int fromRow, int fromCol, CellStoreEnumerator<object> fe, bool atEnd)
        {
            var startCol = fe._startCol;           
            var endRow = fe._endRow;
            var endCol = fe._endCol;
            if (++fromCol > fe._endCol)
            {
                if (endRow <= fromRow) return;
                fromCol = startCol;
                fromRow++;
            }
            int toRow, toCol;
            if (atEnd || fe.Column < 0 || endRow < fe.Row || endCol < fe.Column) 
            {
                toRow = endRow;
                toCol = endCol;
            }
            else
            {
                toRow = fe.Row;
                toCol = fe.Column;
            }

            FormulaRangeAddress fa;
            if(fe._startRow == endRow || startCol==endCol)
            {
                fa = new FormulaRangeAddress() { FromCol = fromCol, FromRow = fromRow, ToCol = toCol, ToRow = toRow };
                rd.Merge(ref fa);
            }
            else if (fromRow < toRow)
            {
                if(fromCol > startCol)
                {
                    fa = new FormulaRangeAddress() { FromCol = fromCol, FromRow = fromRow, ToCol=endCol, ToRow=fromRow};
                    rd.Merge(ref fa);
                    fromRow++;
                }
                if(fromRow < toRow)
                {
                    if(toCol == endCol)
                    {
                        fa = new FormulaRangeAddress() { FromCol = startCol, FromRow = fromRow, ToCol = endCol, ToRow = toRow };
                        rd.Merge(ref fa);
                        return;
                    }
                    fa = new FormulaRangeAddress() { FromCol = startCol, FromRow = fromRow, ToCol = endCol, ToRow = toRow-1 };
                    rd.Merge(ref fa);
                    fromRow = toRow;
                }
                if(fromRow==toRow)
                {
                    fa = new FormulaRangeAddress() { FromCol = startCol, FromRow = toRow, ToCol = toCol, ToRow = toRow };
                    rd.Merge(ref fa);
                }
            }
            else
            {
                fa = new FormulaRangeAddress() { FromCol = fromCol, FromRow = fromRow, ToCol = toCol, ToRow = fromRow };
                rd.Merge(ref fa);
            }
        }

        private static RangeHashset AddAddressToRD(RpnOptimizedDependencyChain depChain, int wsIx)
        {
            if (wsIx < 0) wsIx=-1; //Workboook names
            if (depChain.accessedRanges.TryGetValue(wsIx, out RangeHashset rd) == false)
            {
                rd = new RangeHashset();
                depChain.accessedRanges.Add(wsIx, rd);
            }

            return rd;
        }

        private static void CheckCircularReferences(RpnOptimizedDependencyChain depChain, RpnFormula f, FormulaRangeAddress address, ExcelCalculationOption options)
        {
            if (f._ws == null) return;
            if(f._arrayIndex>=0)
            {
                var sf = f._ws._sharedFormulas[f._arrayIndex];
                var fa = new FormulaRangeAddress(depChain._parsingContext) { FromRow = sf.StartRow, ToRow = sf.EndRow, FromCol = sf.StartCol, ToCol = sf.EndCol, WorksheetIx = f._ws.IndexInList };
                if (fa.CollidesWith(address) != eAddressCollition.No)
                {
                    throw new CircularReferenceException($"Circular reference in Arrayformula: {fa.Address}");
                }
            }
            var wsIx=f._ws?.IndexInList ?? ushort.MaxValue;
            if (address.CollidesWith(wsIx, f._row, f._column))
            {
                var fId = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                HandleCircularReference(depChain, f, options, fId);
            }

            foreach (var sf in depChain._formulaStack)
            {
                wsIx = sf._ws?.IndexInList ?? ushort.MaxValue;
                var toCell = ExcelCellBase.GetCellId(wsIx, sf._row, sf._column);
                if(address.CollidesWith(wsIx, sf._row, sf._column))
                {
                    HandleCircularReference(depChain, f, options, toCell);
                }
            }
        }
        private static void CheckCircularReferences(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options)
        {
            if (f._ws == null) return;

            var cc = depChain._parsingContext.CurrentCell;
            var address = new FormulaRangeAddress() { FromRow = cc.Row, ToRow = cc.Row, FromCol = cc.Column, ToCol = cc.Column };
            foreach (var sf in depChain._formulaStack)
            {
                var sheetId = sf._ws?.IndexInList??ushort.MaxValue;
                if (address.CollidesWith(sheetId, sf._row, sf._column))
                {
                    var toCell = ExcelCellBase.GetCellId(sheetId, sf._row, sf._column);
                    HandleCircularReference(depChain, f, options, toCell);
                }
            }
        }

        private static void HandleCircularReference(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, ulong toCell)
        {
            if (options.AllowCircularReferences)
            {
                //var refFormula = depChain._formulaStack.Peek();
                var fromCell = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                depChain._circularReferences.Add(new CircularReference(fromCell, toCell));
            }
            else
            {
                
                throw new CircularReferenceException($"Circular reference in cell {f.GetAddress()}");
            }
        }

        private static FormulaRangeAddress ExecuteNextToken(RpnOptimizedDependencyChain depChain, RpnFormula f, bool returnAddresses)
        {
            var s = f._expressionStack;
            while (f._tokenIndex < f._tokens.Count)
            {
                var t = f._tokens[f._tokenIndex];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                    case TokenType.Integer:
                    case TokenType.Decimal:
                    case TokenType.StringContent:
                    case TokenType.Array:
                    case TokenType.ParameterVariableDeclaration:
                    case TokenType.ParameterVariable:
                        s.Push(f._expressions[f._tokenIndex]);
                        break;
                    case TokenType.Negator:                        
                        s.Push(s.Pop().Negate());
                        break;
                    case TokenType.CellAddress:
                    case TokenType.ExcelAddress:
						var e = f._expressions[f._tokenIndex];
                        s.Push(e);
                        if(returnAddresses && (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek())==false))
                        {
                           return e.GetAddress();
                        }
                        break;
					case TokenType.FullColumnAddress:
					case TokenType.FullRowAddress:
						var frce = f._expressions[f._tokenIndex];
						s.Push(frce);
                        break;
					case TokenType.NameValue:
                        var ne = (NamedValueExpression)f._expressions[f._tokenIndex];
                        s.Push(ne);
                        if (ne._name != null)
                        {
                            var address = ne.GetAddress();
                            if(address == null)
                            {
                                if (string.IsNullOrEmpty(ne._name?.Formula) == false)
                                {
                                    return null;
                                }
                            }
                            else if (returnAddresses && (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false))
                            {
                                return address;
                            }
                        }
                        break;
                    case TokenType.Comma:
                        if(f._funcStack.Count > 0)
                        {
                            var fexp = f._funcStack.Peek();
                            if(fexp.HandlesVariables && f._expressionStack.Count > 1 && !(f._expressionStack.Peek() is VariableExpression varExp && varExp.IsDeclaration))
                            {
                                var exp1 = f._expressionStack.Pop();
                                var exp2 = f._expressionStack.Pop();
                                f._expressionStack.Push(exp2);
                                f._expressionStack.Push(exp1);
                                if(exp2 is VariableExpression vfe && vfe.IsDeclaration)
                                {
                                    ((VariableFunctionExpression)fexp).AddVariableValue(vfe.Name, exp1.Compile());
                                }
                            }
                            
                            if (f._tokenIndex > 0 && f._tokens[f._tokenIndex - 1].TokenType == TokenType.Comma) //Empty function argument.
                            {
                                //if(fexp._function.HasNormalArguments) fexp._arguments.Add(f._tokenIndex);
                                f._expressionStack.Push(new EmptyExpression());                                
                            }
                            var pi = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos++);
                            if (EnumUtil.HasFlag(pi, FunctionParameterInformation.Condition))
                            {
                                var v = s.Pop().Compile();
                                PushResult(depChain._parsingContext, f, v);
                                fexp._latestConditionValue = GetCondition(v);
                                f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                            }
                            else if (fexp._latestConditionValue==ExpressionCondition.True || fexp._latestConditionValue == ExpressionCondition.False)
                            {
                                pi = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos);
                                if ((pi == FunctionParameterInformation.UseIfConditionIsFalse && fexp._latestConditionValue == ExpressionCondition.True)
                                   ||
                                   (pi == FunctionParameterInformation.UseIfConditionIsTrue && fexp._latestConditionValue == ExpressionCondition.False))
                                {
                                    f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                                }
                            }
                            else if(fexp._latestConditionValue==ExpressionCondition.Error)
                            {
                                f._expressionStack.Push(Expression.Empty);
                                f._expressionStack.Push(Expression.Empty);
                                f._tokenIndex = fexp._endPos-1;
                            }
                        }
                        break;
                    case TokenType.Function:
                        FunctionExpression funcExp;
                        try
                        {
                            if (f._currentFunction == null)
                            {
                                funcExp = f._funcStack.Pop();

                                if (PreExecFunc(depChain, f, funcExp))
                                {
                                    f._currentFunction = funcExp;
                                    f._tokenIndex--; //We should stay on this token when we continue on this formula.
                                    return funcExp._dependencyAddresses.Dequeue();
                                }
                            }
                            else
                            {
                                funcExp = f._currentFunction;
                                if (funcExp._dependencyAddresses.Count > 0)
                                {
                                    f._tokenIndex--; //We should stay on this token when we continue on this formula.
                                    return funcExp._dependencyAddresses.Dequeue();
                                }
                                f._currentFunction = null;
                            }

                            var r = ExecFunc(depChain, f, funcExp);
                            if (r.Address!=null && returnAddresses)
                            {
                                if ((f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false) && r.Address != null)
                                {
                                    return r.Address.Clone();
                                }
                            }
                        }
                        catch
                        {
                            f._expressionStack.Push(ErrorExpression.ValueError);
                        }
                        break;
                    case TokenType.StartFunctionArguments:
                        var fe = (FunctionExpression)f._expressions[f._tokenIndex];
                        if(fe._function==null)  //Function does not exists. Push #NAME?
                        {
                            f._tokenIndex = fe._endPos;
                            f._expressionStack.Push(new ErrorExpression(new CompileResult(eErrorType.Name), depChain._parsingContext));
                            break;
                        }
                        f._funcStack.Push(fe);
                        break;
                    case TokenType.Operator:
                        ApplyOperator(depChain._parsingContext, t, f);
                        break;
                    case TokenType.Percent:
                        ApplyPercent(depChain._parsingContext, f);
                        break;
                    case TokenType.InvalidReference:
                        s.Push(ErrorExpression.RefError);
                        break;
                    case TokenType.ValueDataTypeError:
                        s.Push(ErrorExpression.ValueError);
                        break;
                    case TokenType.NumericError:
                        s.Push(ErrorExpression.NumError);
                        break;
                    case TokenType.NAError:
                        s.Push(ErrorExpression.NaError);
                        break;
                    case TokenType.NameError:
                        s.Push(ErrorExpression.NameError);
                        break;
                }
                f._tokenIndex++;
                if(f._tokenIndex==f._tokens.Count)
                {
					if (s.Count > 0 && s.Peek().Status == ExpressionStatus.IsAddress)
					{
						var cr = s.Peek().Compile();
						return cr.Address;
					}
				}
			}
            return null;
        }

        private static ExpressionCondition GetCondition(CompileResult v)
        {
            if(v.ResultValue is IRangeInfo ri)
            {
                var ret = ExpressionCondition.None;
                for(int r=0;r<ri.Size.NumberOfRows;r++)
                {
                    for (int c = 0; c < ri.Size.NumberOfCols; c++)
                    {
                        var c1 = ConvertUtil.GetValueBool(ri.GetOffset(r, c));
                        if (c1.HasValue)
                        {
                            if (ret == ExpressionCondition.None)
                            {
                                ret = c1.Value ? ExpressionCondition.True : ExpressionCondition.False;
                            }
                            else
                            {
                                var c2 = c1.Value ? ExpressionCondition.True : ExpressionCondition.False;
                                if (c2 != ret)
                                {
                                    return ExpressionCondition.Multi;
                                }
                            }
                        }
                        else
                        {
                            if(ret==ExpressionCondition.None)
                            {
                                ret= ExpressionCondition.Error;
                            }
                            else
                            {
                                return ExpressionCondition.Multi;
                            }
                        }
                    }
                }
                return ret;
            }
            else
            {
                var condition = ConvertUtil.GetValueBool(v.ResultValue);
                if (condition.HasValue)
                {
                    return condition.Value ? ExpressionCondition.True : ExpressionCondition.False;
                }
                return ExpressionCondition.Error;
            }
        }

        private static void ApplyPercent(ParsingContext context, RpnFormula f)
        {
            var e = f._expressionStack.Pop();
            var v=e.Compile().ResultNumeric;
            v /= 100;
            f._expressionStack.Push(new DecimalExpression(new CompileResult(v, DataType.Decimal), context));
        }

        private static bool ShouldIgnoreAddress(FunctionExpression fe)
        {
            if(fe._function.ParametersInfo.HasNormalArguments==false)
            {
                var pi = fe._function.ParametersInfo.GetParameterInfo(fe._argPos);
                return (pi & (FunctionParameterInformation.IgnoreAddress | FunctionParameterInformation.AdjustParameterAddress)) != 0;
            }
            return false;
        }

        private static int GetNextTokenPosFromCondition(RpnFormula f, FunctionExpression fexp)
        {
            if(fexp._argPos < fexp.NumberOfArguments)
            {
                var fe = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos);
                while(fexp._argPos < fexp.NumberOfArguments && (
                    (EnumUtil.HasFlag(fe, FunctionParameterInformation.UseIfConditionIsTrue) && (fexp._latestConditionValue == ExpressionCondition.False || fexp._latestConditionValue == ExpressionCondition.Error)) ||
                    (EnumUtil.HasFlag(fe, FunctionParameterInformation.UseIfConditionIsFalse) && (fexp._latestConditionValue == ExpressionCondition.True || fexp._latestConditionValue == ExpressionCondition.Error))
                    ))
                {
                    fexp._argPos++;
                    //If the argument is empty and it's the last argument it's added in the exec function (first in the GetFunctionArguments method) instead.
                    if (!(f._tokenIndex + 1 < f._tokens.Count && 
                       f._tokens[f._tokenIndex].TokenType == TokenType.Comma && 
                       f._tokens[f._tokenIndex + 1].TokenType == TokenType.Function))
                    {
                        f._expressionStack.Push(Expression.Empty);  //This expression is not used.
                    }
                    fe = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos);
                }
                if(fexp._argPos < fexp.NumberOfArguments)
                {
                    return fexp.GetArgument(fexp._argPos);
                }
                else
                {
                    return fexp._endPos - 1;
                }
            }
            return f._tokenIndex;
        }

#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        private static void ApplyOperator(ParsingContext context, Token opToken, RpnFormula f)
        {
            if (f._expressionStack.Count == 1 && opToken.Value == "=" && f._tokenIndex == f._tokens.Count - 1) 
                return;

            var v1 = f._expressionStack.Pop();
            var v2 = f._expressionStack.Pop();

            
            var c1 = v1.Compile();
            var c2 = v2.Compile();

            if (OperatorsDict.Instance.TryGetValue(opToken.Value, out IOperator op))
            {
                var result = op.Apply(c2, c1, context);
                PushResult(context, f, result);
            }
        }
#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        private static bool PreExecFunc(RpnOptimizedDependencyChain depChain, RpnFormula f, FunctionExpression funcExp)
        {
            IList<CompileResult> args;
            if (_cacheExpressions)
            {
                var cache = depChain.GetCache(f._ws);
                var key = funcExp.GetExpressionKey(f);
                if (string.IsNullOrEmpty(key) || !cache.TryGetValue(key, out funcExp._cachedCompileResult))
                {
                    args = CompileFunctionArguments(f, funcExp);
                    return funcExp.SetArguments(args);
                }
                else
                {
                    //Remove all function arguments from the stack
                    for (int i = 0; i < funcExp.NumberOfArguments; i++)
                    {
                        var si = f._expressionStack.Pop();
                    }
                    funcExp.Status = ExpressionStatus.IsCached;
                }
            }
            else
            {
                args = CompileFunctionArguments(f, funcExp);
                return funcExp.SetArguments(args);
            }                
            return false;
        }

        private static CompileResult ExecFunc(RpnOptimizedDependencyChain depChain, RpnFormula f, FunctionExpression funcExp)
        {
            CompileResult result;
            if (funcExp.Status==ExpressionStatus.IsCached)
            {
                result = funcExp._cachedResult;
            }
            else
            {
                result = funcExp.Compile();
            }
            if(funcExp._function!=null && funcExp._function.ReturnsReference && result.Address!=null)
            {
                f._expressionStack.Push(new RangeExpression(result.Address));
            }
            else
            {
                PushResult(depChain._parsingContext, f, result);
            }
            return result;
        }
        private static void PushResult(ParsingContext context, RpnFormula f, CompileResult result)
        {
            switch (result.DataType)
            {
                case DataType.Boolean:
                    f._expressionStack.Push(new BooleanExpression(result, context));
                    break;
                case DataType.Integer:
                    f._expressionStack.Push(new DecimalExpression(result, context));
                    break;
                case DataType.Decimal:
                case DataType.Date:
                case DataType.Time:
                    f._expressionStack.Push(new DecimalExpression(result, context));
                    break;
                case DataType.String:
                    f._expressionStack.Push(new StringExpression(result, context));
                    break;
                case DataType.ExcelError:
                    f._expressionStack.Push(new ErrorExpression(result, context));
                    break;
                case DataType.ExcelRange:
                    f._expressionStack.Push(new RangeExpression(result, context));
                    break;
                case DataType.Empty:
                    f._expressionStack.Push(Expression.Empty);
                    break;
                default:
                    //throw new InvalidOperationException($"Unhandled compile result for data type {result.DataType}");
                    f._expressionStack.Push(ErrorExpression.ValueError);
                    break;
            }
        }


        private static IList<CompileResult> CompileFunctionArguments(RpnFormula f, FunctionExpression func)
        {
            var list = new List<CompileResult>();
            if (f._tokenIndex > 0 && f._tokens[f._tokenIndex - 1].TokenType == TokenType.Comma) //Empty function argument.
            {
                f._expressionStack.Push(new EmptyExpression());
            }
            var s = f._expressionStack;
            for(int i=0;i<func.NumberOfArguments && s.Count > 0;i++)
            {
                var si = s.Pop();
                if(si.ExpressionType!=ExpressionType.Empty)
                {
                    si.Status |= ExpressionStatus.FunctionArgument;
                }
                list.Insert(0, si.Compile());
            }
            return list;
        }

        private static bool GetProcessedAddress(RpnOptimizedDependencyChain depChain, ref FormulaRangeAddress address)
        {
            if (depChain.accessedRanges.TryGetValue(address.WorksheetIx, out RangeHashset wsRd) == false)
            {
                wsRd = new RangeHashset();
                depChain.accessedRanges.Add(address.WorksheetIx, wsRd);
            }
            return wsRd.Merge(ref address);
        }
        private static bool GetProcessedAddress(RpnOptimizedDependencyChain depChain, int wsIndex, int row, int col)
        {
            if (depChain.accessedRanges.TryGetValue(wsIndex, out RangeHashset wsRd) == false)
            {
                wsRd = new RangeHashset();
                depChain.accessedRanges.Add(wsIndex, wsRd);
            }
            return wsRd.Merge(row, col);
        }
    }
}