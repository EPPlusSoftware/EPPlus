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
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.FormulaParsing.DependencyChain;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml.Linq;
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
#if !NET35
                options.CancellationToken.ThrowIfCancellationRequested();
#endif
                if (ws.IsChartSheet == false)
                {
                    ExecuteChain(depChain, ws.Cells, options, true);
                    ExecuteChain(depChain, ws.Names, options, true);
                }
            }
            ExecuteChain(depChain, wb.Names, options, true);
            
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
                ws = wb.GetWorksheetByIndexInList(cell.WorksheetIx);
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
            return CalculateFormulaChain(depChain, formula, options, false).Result;
        }

        private static void ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelRangeBase range, ExcelCalculationOption options, bool writeToCell)
        {
            var ws = range.Worksheet;
            RpnFormula f = null;
#if !NET35
            var ct = options.CancellationToken; // Cache locally — avoids property lookup in hot loop
#endif
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {
#if !NET35
                ct.ThrowIfCancellationRequested(); // P0 – per cell
#endif
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
                {
                    try
                    {
                        if (GetFormula(depChain, ws, fs.Row, fs.Column, fs.Value, ref f))
                        {
                            CalculateFormulaChain(depChain, f, options, writeToCell);
                        }
                    }
#if !NET35
                    catch (OperationCanceledException)
                    {
                        throw; // Must propagate — do not swallow
                    }
#endif
                    catch (CircularReferenceException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        if (writeToCell)
                        {
                            SetAndReturnValueError(depChain, ex, f);
                        }
                    }
                }
            }

            if (depChain.HasAnyArrayFormula) //Array formulas has been update. Check if we need to set the array flag on any calculated tables on intersecting tables.
            {
                UpdateTableArrayFlag(range);
            }
        }

        private static void UpdateTableArrayFlag(ExcelRangeBase range)
        {
            //Check table formulas that needs the array flag updated for the columns formulas.
            foreach (var table in range.Worksheet.Tables)
            {
                if (table.Address.Collide(range) != eAddressCollition.No)
                {
                    foreach (var c in table.Columns)
                    {
                        if (string.IsNullOrEmpty(c.CalculatedColumnFormula) == false)
                        {
                            var ca = c.DataAddress;
                            if (ca.Collide(range) != eAddressCollition.No)
                            {
                                var ma = ca.Intersect(range);
                                c.IsCalculatedFormulaArray = IsCFArray(range.Worksheet, ma);
                            }
                        }
                    }
                }
            }
        }

        private static bool IsCFArray(ExcelWorksheet ws, ExcelAddressBase ma)
        {
            for (int row = ma._fromRow; row <= ma._toRow; row++)
            {
                var f = (CellFlags)ws._flags.GetValue(row, ma._fromCol);
                if ((f & CellFlags.ArrayFormula) != 0)
                {
                    return true;
                }
            }
            return false;
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
            if (depChain._parsingContext.CurrentWorksheet != null)
            {
                if (cc.Column > 0)
                {
                    depChain._parsingContext.CurrentWorksheet.SetValueInner(cc.Row, cc.Column, ret);
                }
                else if (cc.Row >= 0 && cc.Row < depChain._parsingContext.CurrentWorksheet.Names.Count)
                {
                    depChain._parsingContext.CurrentWorksheet.Names[cc.Row].Value = ret;
                }
            }
            else if (cc.Column == 0 && cc.Row >= 0 && cc.Row < depChain._parsingContext.Package.Workbook.Names.Count)
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
#if !NET35
            catch (OperationCanceledException)
            {
                throw; // Must propagate
            }
#endif
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
            if (ws != null && ws.IsDisposed) return;
            var wsIx = ws == null ? -1 : ws.IndexInList;
            depChain._parsingContext.CurrentCell = new FormulaCellAddress(wsIx, name.Index, 0);
            var id = ExcelCellBase.GetCellId(wsIx, name.Index, 0);
            if (depChain.processedCells.Contains(id) == false)
            {
                if (string.IsNullOrEmpty(name.NameFormula) == false)
                {
                    var f = GetNameFormula(depChain, ws, depChain._parsingContext.ExcelDataProvider.GetName(name), 1, 1);
                    CalculateFormulaChain(depChain, f, options, writeToCell);
                }
                else if (ws != null && name.IsValidRowCol())
                {
                    ExecuteChain(depChain, name, options, writeToCell);
                }
                else if (name.NameValue != null)
                {
                    name.Value = name.NameValue;
                }
                else
                {
                    name.Value = ErrorValues.RefError;
                }
            }
        }

        private static object ExecuteChain(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, string formula, FormulaCellAddress cell, ExcelCalculationOption options, bool writeToCell)
        {
            try
            {
                var f = new RpnFormula(ws, cell.Row, cell.Column);
                depChain._parsingContext.CurrentCell = new FormulaCellAddress(ws?.Index??-1, -1, 0);
                f.SetFormula(formula, depChain);
                return CalculateFormulaChain(depChain, f, options, writeToCell).Result;
            }
#if !NET35
            catch (OperationCanceledException)
            {
                throw; // Must propagate
            }
#endif
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
                return CalculateFormulaChain(depChain, f, options, writeToCell).Result;
            }
#if !NET35
            catch (OperationCanceledException)
            {
                throw; // Must propagate
            }
#endif
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
        private static bool GetFormula(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, int row, int column, object value, ref RpnFormula f)
        {

            if (value == null) return false;
            if (value is int ix)
            {
                var sf = ws._sharedFormulas[ix];
                if (sf.FormulaType == FormulaType.Array)
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
                        f._flags |= FormulaFlags.IsDynamic;
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
            CheckAndClearRichData(f);
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
            else if (f.Type == RpnFormulaType.NameFormula)
            {
                var cc = ((RpnNameFormula)f).CurrentCell;
                if (cc.Row == 0) cc = new FormulaCellAddress(f._ws == null ? -1 : f._ws.IndexInList, f._row, f._column); //Not set, set to the name.
                depChain._parsingContext.CurrentCell = cc;
            }
        }
        private static RpnFormula GetNameFormula(RpnOptimizedDependencyChain depChain, ExcelWorksheet ws, INameInfo name, int cellRow, int cellCol)
        {
            ExcelCellBase.SplitCellId(name.Id, out int wsIx, out int row, out int col);
            if (name.wsIx >= 0 && ws == null && depChain._parsingContext.Package.Workbook.Worksheets.Count > name.wsIx)
            {
                ws = depChain._parsingContext.Package.Workbook.Worksheets[name.wsIx];
            }
            var f = new RpnNameFormula(ws, row, col, depChain._parsingContext.CurrentCell);
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

        internal static CompileResult ExecutePartialFormula(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, bool writeToCell)
        {
            return CalculateFormulaChain(depChain, f, options, writeToCell);
        }

        private static CompileResult CalculateFormulaChain(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, bool writeToCell, int depChainPos=-1)
        {
            FormulaRangeAddress[] addresses;
            RangeHashset rd = AddOrGetRDFromWsIx(depChain, f._ws == null ? -1 : f._ws.IndexInList);
            object v = null;
            bool hasLogger = depChain._parsingContext.Parser.Logger != null;
            var followChain = options.FollowDependencyChain;
            if (depChainPos == -1)
            {
                rd?.Merge(f._row, f._column);
                depChain.StartOfChain();
            }

        ExecuteFormula:
            try
            {
#if !NET35
                options.CancellationToken.ThrowIfCancellationRequested(); // P0 – per dependency step
#endif
                SetCurrentCell(depChain, f);
                var ws = f._ws;

                if (f._tokenIndex < f._tokens.Count)
                {
                    addresses = ExecuteNextToken(depChain, f, followChain);
                    if (f._tokenIndex < f._tokens.Count)
                    {
                        if (addresses == null && f._expressions.ContainsKey(f._tokenIndex) && f._expressions[f._tokenIndex].ExpressionType == ExpressionType.NameValue)
                        {
                            var ne = f._expressions[f._tokenIndex] as NamedValueExpression;
                            if (ne._externalReferenceIx < 1)
                            {
                                rd = AddOrGetRDFromWsIx(depChain, ne._worksheetIx);

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

                        if (addresses == null)
                        {
                            addresses = f._expressions[f._tokenIndex].GetAddress();
                        }
                        depChain.AddFormulaToChain(f, addresses);
                        if (GetAddressesToFollow(depChain, f, options, ref addresses, ref rd, ref ws))
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

                if (cr != null && f.IsLambda == false &&  (writeToCell || depChain._formulaStack.Count > 0))  // If calculating single cell via the FormulaParser.Parse method we should not write to the cells
                {
                    if (f._ws != null)
                    {
                        rd = AddOrGetRDFromWsIx(depChain, f._ws.IndexInList);
                    }
                    SetValueToWorkbook(depChain, f, rd, cr, options, ref depChainPos);


                    //We are in a dirty cell recalculation and have a new position in the chain.
                    //We should return to the caller and let it continue from the new position in the chain.
                    //We use this technique to avoid stack overflow exceptions when recalculating dirty cells with long dependency chains.
                    if (depChain._recalculateDirtyCellsNewPosition)  
                    {
                        depChain._recalculateDirtyCellsNewPosition = false;
                        return cr;
                    }
                }

                if (hasLogger)
                {
                    depChain._parsingContext.Parser.Logger.Log($"Set value in Cell\t{f.GetAddress()}\t{cr.ResultValue}\t{cr.DataType}");
                }

                if (depChain._formulaStack.Count > f._lambdaFormulaStackCount)
                {
                    f = depChain._formulaStack.Pop();
                    if (f._formulaEnumerator == null)
                    {
                        f._tokenIndex++;
                        goto ExecuteFormula;
                    }

                    rd = AddOrGetRDFromWsIx(depChain, f._enumeratorWorksheetIx);
                    goto NextFormula;
                }
                return cr;
            FollowChain:
                if (addresses.Length == 0)
                {
                    f._tokenIndex++;
                    goto ExecuteFormula;
                }

                var firstAddress = addresses.FirstOrDefault();
                ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(firstAddress.WorksheetIx);
                if (ws == null)
                {
                    f._tokenIndex++;
                    goto ExecuteFormula;
                }
                if (addresses.Length == 1 && addresses[0].IsSingleCell)
                {
                    if (depChain.processedCells.Contains(ExcelCellBase.GetCellId(ws?.IndexInList ?? ushort.MaxValue, firstAddress.FromRow, firstAddress.FromCol)) == false)
                    {

                        rd?.Merge(firstAddress.FromRow, firstAddress.FromCol);

                        if (ws._formulas.Exists(firstAddress.FromRow, firstAddress.FromCol, ref v) && v != null)
                        {
                            depChain._formulaStack.Push(f);
                            GetFormula(depChain, ws, firstAddress.FromRow, firstAddress.FromCol, v, ref f);
                            goto ExecuteFormula;
                        }
                    }
                    f._tokenIndex++;

                    goto ExecuteFormula;
                }
                else
                {
                    f._enumeratorWorksheetIx = ws.IndexInList;
                    f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, addresses);
                }
            NextFormula:
                var fe = f._formulaEnumerator;
                var row = fe.Row;
                var col = fe.Column < 0 ? fe._startCol : fe.Column;
                var rPos = fe.RangePos;
                if (fe.Next())
                {
                    if (fe.Value == null || depChain.processedCells.Contains(ExcelCellBase.GetCellId(f._enumeratorWorksheetIx, fe.Row, fe.Column)))
                    {
                        MergeToRd(rd, row, col, rPos, fe, false);
                        goto NextFormula;
                    }

                    depChain._formulaStack.Push(f);
                    MergeToRd(rd, row, col, rPos, fe, false);
                    if (GetFormula(depChain, ws, fe.Row, fe.Column, fe.Value, ref f))
                    {
                        goto ExecuteFormula;
                    }
                    else
                    {
                        goto NextFormula;
                    }
                }
                MergeToRd(rd, row, col, rPos, fe, true);

                f._formulaEnumerator = null;
                f._tokenIndex++;

                goto ExecuteFormula;
            }
#if !NET35
            catch (OperationCanceledException)
            {
                throw; // Must propagate
            }
#endif
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

                f._tokenIndex = f._tokens.Count - 1;
                if (depChain._formulaStack.Count > 0)
                {
                    f = depChain._formulaStack.Pop();
                    goto ExecuteFormula;
                }
                //goto CheckFormulaStack;
                return new CompileResult(errValue, DataType.ExcelError);
            }

        }


        private static bool GetAddressesToFollow(RpnOptimizedDependencyChain depChain, RpnFormula f, ExcelCalculationOption options, ref FormulaRangeAddress[] addresses, ref RangeHashset rd, ref ExcelWorksheet ws)
        {
            var hasAddress = false;
            var needsClean = false;
            for (int i = 0; i < addresses.Length; i++)
            {
                if (addresses[i].FromRow<1 || addresses[i].FromCol<1)
                {
                    addresses[i] = null;
                    needsClean = true;
                    continue;
                }
                var address = addresses[i].Clone();
                if (address.ExternalReferenceIx > 0) //We don't follow dep chain into external references.
                {
                    addresses = null;
                    return false;
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
                else if (address?.WorksheetIx >= 0 && ws?.IndexInList != address.WorksheetIx)
                {
                    ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(address.WorksheetIx);
                }
                if (ws == null) return false;
                rd = AddOrGetRDFromWsIx(depChain, ws.IndexInList);

                if (rd.Exists(address) || address.CollidesWith(ws.IndexInList, f._row, f._column))
                {
                    CheckCircularReferences(depChain, f, address, options);
                }

                if (rd.ExistsGetSpill(ref address))
                {
                    addresses[i] = address;
                    hasAddress = true;
                }
                else
                {

                    addresses[i] = null;
                    needsClean = true;
                }
            }
            if (needsClean)
            {
                addresses = addresses.Where(x => x != null).ToArray();
            }
            return hasAddress;
        }

        private static void CheckAndClearRichData(RpnFormula f)
        {
            var ws = f._ws;
            if (ws == null) return;
            var md = f._ws._metadataStore.GetValue(f._row, f._column);
            if (md.vm > 0u)
            {
                var mdb = ws.Workbook.Metadata.Db.ValueMetadata.Get(md.vm);
                if (mdb != null)
                {
                    var rv = f._ws._richDataStore.GetRichValue(md.vm);
                    if (rv != null && rv.Structure.Type != "_webimage")
                    {
                        mdb.DeleteMe();
                    }
                    else
                    {
                        return;
                    }
                }
            }
            if (md.cm > 0u)
            {
                var metadata = ws.Workbook.Metadata;
                if (!metadata.DynamicArrayTypeId.HasValue || md.cm != metadata.DynamicArrayTypeId.Value)
                {
                    var cdb = metadata.Db.CellMetadata.Get(md.cm);
                    if (cdb != null)
                    {
                        cdb.DeleteMe();
                    }
                }
            }
            f._ws._metadataStore.Clear(f._row, f._column, 1, 1);
        }
        private static void SetValueToWorkbook(RpnOptimizedDependencyChain depChain, RpnFormula f, RangeHashset rd, CompileResult cr, ExcelCalculationOption options, ref int insertDepChainPos)
        {
            if(cr.DataType == DataType.LambdaCalculation)
            {
                cr = CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
            }
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
                        if ((cr.DataType == DataType.ExcelRange && ((IRangeInfo)cr.Result).Address.IsSingleCell == false)) //A range. When we add support for dynamic array formulas we will alter this.
                        {
                            var ri = (IRangeInfo)cr.Result;
                            if (f._arrayIndex >= 0 && (f._flags & FormulaFlags.IsDynamic) == 0) //A legacy array formula, Fill the referenced range.
                            {
                                ArrayFormulaOutput.FillArrayFromRangeInfo(f, ri, rd, depChain);
                                depChain.HasAnyArrayFormula = true;
                            }
                            else
                            {
                                if (f.CanBeDynamicArray) //Create a dynamic array formula if allowed. 
                                {
                                    //Add dynamic array formula support here.
                                    var dirtyRange = ArrayFormulaOutput.FillDynamicArrayFromRangeInfo(f, ri, rd, depChain);

                                    if (dirtyRange != null && dirtyRange.Length > 0)
                                    {
                                        RecalculateDirtyCells(dirtyRange, depChain, rd, options);
                                    }
                                }
                                else //Set implicit intersection
                                {
                                    var icr = ImplicitIntersectionUtil.GetResult(ri, f._row, f._column, depChain._parsingContext);
                                    f._ws.SetValueInner(f._row, f._column, icr.ResultValue ?? 0D);
                                }
                            }
                        }
                        else if ((cr.ResultType == CompileResultType.DynamicArray ||
                                 cr.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic ||
                                 (f._flags & FormulaFlags.IsAlwaysDynamic) == FormulaFlags.IsAlwaysDynamic) &&
                                 f.CanBeDynamicArray)
                        {

                            var dirtyRange = ArrayFormulaOutput.FillDynamicArraySingleValue(f, cr, rd, depChain);

                            if (dirtyRange != null && dirtyRange.Length > 0)
                            {
                                RecalculateDirtyCells(dirtyRange, depChain, rd, options);
                            }

                            depChain.HasAnyArrayFormula = true;
                        }
                        else if (cr.ResultType == CompileResultType.LocalImage)
                        {
                            var picManager = new CellPicturesManager(f._ws);
                            var pic = cr.Result as ExcelCellPicture;
                            picManager.SetCellPicture(f._row, f._column, pic.GetImageBytes(), pic.AltText, CalcOrigins.Reference);
                        }
                        else if (cr.ResultType == CompileResultType.WebImage)
                        {
                            var pic = cr.Result as ExcelCellPicture;
                            if (pic.IsReferenceTo(f._ws.Name, f._row, f._column))
                            {
                                var picManager = new CellPicturesManager(f._ws);
                                picManager.SetWebPicture(f._row, f._column, pic.ExternalAddress, pic.GetImageBytes(), pic.AltText, CalcOrigins.Reference);
                            }
                        }
                        else
                        {
                            if (f._arrayIndex != -1)
                            {
                                if((f._flags & FormulaFlags.IsDynamic)== FormulaFlags.IsDynamic)
                                {
                                    var dirtyRange = ArrayFormulaOutput.FillDynamicArraySingleValue(f, cr, rd, depChain);
                                    if (dirtyRange != null && dirtyRange.Length > 0)
                                    {
                                        RecalculateDirtyCells(dirtyRange, depChain, rd, options);
                                    }
                                    depChain.HasAnyArrayFormula = true;

                                }
                                else
                                {
                                    var sf = f._ws._sharedFormulas[f._arrayIndex];
                                    f._ws.SetValueInner(sf.StartRow, sf.StartCol, sf.EndRow, sf.EndCol, cr.ResultValue ?? 0D);
                                }
                            }
                            else
                            {
                                var dVal = cr.ResultValue;
                                if (cr.DataType == DataType.Decimal && dVal != null && dVal is double dbl && dbl == 0d)
                                {
                                    // this is to avoid "-0" results from the ToString method.
                                    dVal = 0d;
                                }
                                f._ws.SetValueInner(f._row, f._column, dVal ?? 0D);
                            }
                        }
                    }
                }

                if (insertDepChainPos<0)
                {
                    depChain.DependencyChain.Add(f.CellId);
                }
                else if(depChain._formulaStack.Count > 0)
                {
                    depChain.DependencyChain.Insert(++insertDepChainPos, f.CellId);
                }
            }
        }
        private static void RecalculateDirtyCells(SimpleAddress[] dirtyRange, RpnOptimizedDependencyChain depChain, RangeHashset rd, ExcelCalculationOption options)
        {
            if (options.FollowDependencyChain == false)
            {
                return; //EPPlus will not recalculate dirty cells while recalculating other dirty cells to avoid stack overflow.
            }
            if(depChain._recalculateDirtyCellsIterations > ExcelCalculationOption.MaxArrayFormulaRecalculationIterations)
            {
                throw(new DynamicArrayMaxIterationsException($"Too many iterations when recalculating dirty dynamic array formula ranges. EPPlus currently limit this value to {ExcelCalculationOption.MaxArrayFormulaRecalculationIterations} iterations"));
            }
            var dirtyCells = dirtyRange.ToList();
            if (depChain.FormulaRangeReferences.ContainsKey(depChain._parsingContext.CurrentWorksheet.IndexInList))
            {
                var qt = depChain.FormulaRangeReferences[depChain._parsingContext.CurrentWorksheet.IndexInList];
                int fromIx = int.MaxValue;
                //int toIx = int.MinValue;
                foreach (var a in dirtyRange)
                {
                    var ir = qt.GetIntersectingRangeItems(new QuadRange(a.FromRow, a.FromCol, a.ToRow, a.ToCol));
                    if (ir.Count > 0)
                    {
                        depChain._parsingContext.RangeCriteriaCache.Clear();
                        foreach (var r in ir.Select(x => x.Value).Distinct())
                        {
                            ExcelAddressBase.SplitCellId(r, out int sheet, out int row, out int col);

                            var f=depChain._formulaStack.FirstOrDefault(x => x.CellId == r);
                            if (f!=null)
                            {
                                foreach (var fs in depChain._formulaStack)
                                {
                                   fs.Reset(depChain);
                                }
                                continue;
                            }
                            
                            var ix = depChain.DependencyChain.IndexOf(r);
                            if (ix < 0) continue;
                            if (ix < fromIx)
                            {
                                fromIx = ix;
                            }

                            if(fromIx==0)
                            {
                                break;
                            }
                        }
                    }
                    if (fromIx == 0)
                    {
                        break;
                    }
                }

                if(depChain._recalculateDirtyCellsPosition >= 0)
                {
                    depChain._recalculateDirtyCellsIterations++;
                    depChain._recalculateDirtyCellsPosition = fromIx;
                    depChain._recalculateDirtyCellsNewPosition = true;
                    return;
                }

                if (fromIx != int.MaxValue)
                {
                    var dcCount = depChain.DependencyChain.Count;
                    var cc = depChain._parsingContext.CurrentCell;                    

                    for (int i = fromIx; i < dcCount; i++)
                    {
                        ExcelCellBase.SplitCellId(depChain.DependencyChain[i], out int sheetId, out int row, out int col);
                        RpnFormula f = null;
                        var ws = depChain._parsingContext.Package.Workbook.GetWorksheetByIndexInList(sheetId);
                        var v = ws._formulas.GetValue(row, col);

                        if (GetFormula(depChain, ws, row, col, v, ref f))
                        {
                            f.ClearCache(depChain);
                            depChain._recalculateDirtyCellsPosition=i;
                            CalculateFormulaChain(depChain, f, options, true, i);

                            if (depChain.DependencyChain.Count > dcCount)
                            {
                                i += depChain.DependencyChain.Count - dcCount;
                                dcCount = depChain.DependencyChain.Count;
                            }

                            if (depChain._recalculateDirtyCellsPosition < i)
                            {
                                i = depChain._recalculateDirtyCellsPosition - 1; //We have a new position to recalculate from, set i to one before that because of the i++ in the for loop.
                            }
                            else
                            {
                                depChain._recalculateDirtyCellsIterations--;
                            }
                        }
                    }
                    depChain._recalculateDirtyCellsIterations = 0;
                    depChain._recalculateDirtyCellsPosition = -1;
                }
            }
        }
        private static void MergeToRd(RangeHashset rd, int fromRow, int fromCol, int rangePos, CellStoreEnumerator<object> fe, bool atEnd)
        {

            if (rangePos < fe.RangePos)
            {
                var a = fe.Ranges[rangePos];
                if (fromCol < 1) fromCol = 1;
                MergeAddressToRd(rd, fe, fromRow, fromCol, a.ToRow, a.ToCol, a.ToRow, a.ToCol);
                for (int i = rangePos; i < fe.RangePos - 1; i++)
                {
                    a = fe.Ranges[i];
                    MergeAddressToRd(rd, fe, a.FromRow, a.FromCol, a.ToRow, a.ToCol, a.ToRow, a.ToCol);
                }
                fromRow = fe._startRow;
                fromCol = fe._startCol;
            }
            else
            {
                if (fromCol > fe._endCol)
                {
                    if (fe._endRow <= fromRow) return;
                    fromCol = fe._startCol;
                    fromRow++;
                }
            }

            var endRow = fe._endRow;
            var endCol = fe._endCol;
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
            MergeAddressToRd(rd, fe, fromRow, fromCol, toRow, toCol, endRow, endCol);
        }

        private static void MergeAddressToRd(RangeHashset rd, CellStoreEnumerator<object> fe, int fromRow, int fromCol, int toRow, int toCol, int endRow, int endCol)
        {
            var startCol = fe._startCol;
            FormulaRangeAddress fa;
            if (fe._startRow == endRow || startCol == endCol)
            {
                fa = new FormulaRangeAddress() { FromCol = fromCol, FromRow = fromRow, ToCol = toCol, ToRow = toRow };
                rd.Merge(ref fa);
            }
            else if (fromRow < toRow)
            {
                if (fromCol > startCol)
                {
                    fa = new FormulaRangeAddress() { FromCol = fromCol, FromRow = fromRow, ToCol = endCol, ToRow = fromRow };
                    rd.Merge(ref fa);
                    fromRow++;
                }
                if (fromRow < toRow)
                {
                    if (toCol == endCol)
                    {
                        fa = new FormulaRangeAddress() { FromCol = startCol, FromRow = fromRow, ToCol = endCol, ToRow = toRow };
                        rd.Merge(ref fa);
                        return;
                    }
                    fa = new FormulaRangeAddress() { FromCol = startCol, FromRow = fromRow, ToCol = endCol, ToRow = toRow - 1 };
                    rd.Merge(ref fa);
                    fromRow = toRow;
                }
                if (fromRow == toRow)
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

        private static RangeHashset AddOrGetRDFromWsIx(RpnOptimizedDependencyChain depChain, int wsIx)
        {
            if (wsIx < 0) wsIx = -1; //Workboook names
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
            var wsIx = f._ws?.IndexInList ?? ushort.MaxValue;
            if (f._arrayIndex>=0)
            {
                var sf = f._ws._sharedFormulas[f._arrayIndex];
                var fa = new FormulaRangeAddress(depChain._parsingContext) { FromRow = sf.StartRow, ToRow = sf.EndRow, FromCol = sf.StartCol, ToCol = sf.EndCol, WorksheetIx = f._ws.IndexInList };
                if (fa.CollidesWith(address) != eAddressCollition.No)
                {
                    if(!options.AllowCircularReferences)
                    {
                        throw new CircularReferenceException($"Circular reference in array formula: {fa.Address}");
                    }
                    else
                    {
                        var toCell = ExcelCellBase.GetCellId(wsIx, sf.StartRow, sf.StartCol);
                        var fromCell = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                        depChain._circularReferences.Add(new CircularReference(fromCell, toCell));
                    }
                }
            }
            if (address.CollidesWith(wsIx, f._row, f._column))
            {
                var fId = ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column);
                HandleCircularReference(depChain, f, options, fId);
            }

            foreach (var sf in depChain._formulaStack)
            {
                wsIx = sf._ws?.IndexInList ?? ushort.MaxValue;
                var toCell = ExcelCellBase.GetCellId(wsIx, sf._row, sf._column);
                if (address.CollidesWith(wsIx, sf._row, sf._column))
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
                var sheetId = sf._ws?.IndexInList ?? ushort.MaxValue;
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
                var cr = new CircularReference(ExcelCellBase.GetCellId(f._ws.IndexInList, f._row, f._column), toCell);
                if (depChain._circularReferences.Contains(cr)==false)
                {
                    depChain._circularReferences.Add(cr);
                }
            }
            else
            {
                throw new CircularReferenceException($"Circular reference in cell {f.GetAddress()}");
            }
        }

        private static bool IsSingleAddress(RpnFormula f)
        {
            var t = f._tokenIndex + 1;
            while (t < f._tokens.Count && f._tokens[t].TokenTypeIsAddressToken)
            {
                if (f._tokens[t].TokenType == TokenType.Operator && f._tokens[t].Value == ":")
                {
                    return false;
                }
                t++;
            }
            return true;
        }

        private static FormulaRangeAddress[] ExecuteNextToken(RpnOptimizedDependencyChain depChain, RpnFormula f, bool returnAddresses)
        {
            FormulaRangeAddress[] addresses;
            var s = f._expressionStack;
            while (f._tokenIndex < f._tokens.Count)
            {
                if (f.HasLambdaToken(f._tokenIndex))
                {
                    f._tokenIndex++;
                    continue;
                }
                if(LambdaExpressionFunctions.LastExpressionIsLambdaCalculation(s, out LambdaCalculationExpression lce))
                {
                    LambdaExpressionFunctions.PreProcessLambdaCalculation(s, f, lce);
                }
                var leStackPos = f.GetCurrentLambdaExpressionStackPosition();
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
                    case TokenType.CommaLambda:
                    case TokenType.EmptyArgument:
                    case TokenType.EtaReducedLambda:
                        if(leStackPos != null)
                        {
                            var exp = f._expressions[f._tokenIndex];
                            if (exp.ExpressionType == ExpressionType.LambdaVariableDeclaration || exp.ExpressionType == ExpressionType.Variable)
                            {
                                exp.Status |= ExpressionStatus.IsLambdaVariableDeclaration;
                            }
                            if (t.TokenType == TokenType.EtaReducedLambda)
                            {
                                ((FunctionExpression)f._expressions[f._tokenIndex]).SetRpnFormula(f);
                            }

                            var cr = exp.Compile();
                            if (cr.DataType == DataType.LambdaTokens)
                            {
                                s.Push(f._expressions[f._tokenIndex]);
                            }
                            else if(cr.IsVariableResult && f.FunctionStack.Count > 0 && f.FunctionStack.Peek().IsLambda)
                            {
                                s.Push(f._expressions[f._tokenIndex]);
                            }
                            else if(
                                f._expressionStack.Peek() is LambdaCalculationExpression lce2 && lce2.ArgumentCollectionStarted &&
                                cr.DataType != DataType.LambdaVariableDeclaration 
                                && f.LambdaSettings.LambdaArgsAdded.Count > 0 
                                && f.LambdaSettings.LambdaArgsAdded.Peek() < f.GetNumberOfLambdaVariables())
                            {
                                leStackPos.Expression.SetVariable(f.LambdaSettings.LambdaArgsAdded.Peek(), cr.Result, cr.DataType, cr.Address);
                                var nLambdaArgsAdded = f.LambdaSettings.LambdaArgsAdded.Pop();
                                f.LambdaSettings.LambdaArgsAdded.Push(++nLambdaArgsAdded);
                            }
                            else
                            {
                                s.Push(f._expressions[f._tokenIndex]);
                            }
                        }
                        else
                        {
                            s.Push(f._expressions[f._tokenIndex]);
                        }
                        
                        break;
                    case TokenType.Negator:
                        s.Push(s.Pop().Negate());
                        break;
                    case TokenType.CellAddress:
                    case TokenType.ExcelAddress:
                    case TokenType.FullColumnAddress:
                    case TokenType.FullRowAddress:
                        var e = f._expressions[f._tokenIndex];
                        s.Push(e);
                        var localReturnAddress = true;
                        if (leStackPos != null)
                        {
                            var addVariable = true;
                            if(t.TokenType == TokenType.CellAddress)
                            {
                                var tIx = f._tokenIndex + 1;
                                while(tIx < f._tokens.Count - 1 && f._tokens[tIx].TokenType == TokenType.CellAddress)
                                {
                                    addVariable = false;
                                    localReturnAddress = false;
                                    s.Push(f._expressions[tIx]);
                                    f._tokenIndex++;
                                    tIx++;
                                }
                            }
                            if(addVariable && f.LambdaSettings.LambdaArgsAdded.Count > 0)
                            {
                                var rangeCr = e.Compile();
                                leStackPos.Expression.SetVariable(f.LambdaSettings.LambdaArgsAdded.Peek(), rangeCr.ResultValue, rangeCr.DataType, rangeCr.Address);
                                var nLambdaArgsAdded = f.LambdaSettings.LambdaArgsAdded.Pop();
                                f.LambdaSettings.LambdaArgsAdded.Push(++nLambdaArgsAdded);
                            }
                        }
                        if (localReturnAddress && returnAddresses && (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false))
                        {
                            if(f._tokenIndex + 1 < f._tokens.Count)
                            {
                                var ix = f._tokenIndex + 1;
                                var nt = f._tokens[ix];
                                while(f._tokenIndex + 1 < f._tokens.Count && nt.TokenType == TokenType.CellAddress)
                                {
                                    nt = f._tokens[++ix];
                                }
                                if (nt.TokenType == TokenType.Operator && nt.Value == ":")
                                {
                                    f._tokenIndex++;
                                    continue;
                                }
                            }
                            if (t.TokenType == TokenType.CellAddress || t.TokenType == TokenType.ExcelAddress)  //Full column and full row addresses will be returned when processing the : operator.
                            {
                                return e.GetAddress();
                            }
                        }
                        break;
                    case TokenType.NameValue:
                        var ne = (NamedValueExpression)f._expressions[f._tokenIndex];
                        s.Push(ne);
                        if (ne._name != null)
                        {
                            var nameAddress = ne.GetAddress();
                            if (nameAddress == null)
                            {
                                if (returnAddresses && string.IsNullOrEmpty(ne._name?.Formula) == false)
                                {
                                    return null;
                                }
                            }
                            else if (returnAddresses && (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false))
                            {
                                if (IsSingleAddress(f))
                                {
                                    foreach(var a in nameAddress)
                                    {
                                        return GetCriteriaRange(depChain._parsingContext, f, a);
                                    }
                                }
                            }
                        }
                        break;
                    case TokenType.Comma:
                        if(f.ShouldInvokeLambda(s))
                        {
                            CompileResult lambdaResult = LambdaInvoker.InvokeLambdaFunction(depChain, f);
                            if (lambdaResult != null)
                                PushResult(depChain._parsingContext, f, lambdaResult);
                        }
                        else if(f._funcStack.Count > 0)
                        {
                            var fexp = f._funcStack.Peek();
                            if(fexp.IsLet && f.CurrentLambdaArgsAdded == 0 && f._expressionStack.Count > 1 && !(f._expressionStack.First() is VariableExpression varExp && varExp.IsDeclaration))
                            {
                                var exp1 = f._expressionStack.Pop();
                                var exp2 = f._expressionStack.Peek();
                                f._expressionStack.Push(exp1);
                                if (exp2 is VariableExpression vfe && vfe.IsDeclaration)
                                {
                                    ((VariableFunctionExpression)fexp).AddVariableValue(vfe.Name, exp1.Compile());
                                    f._expressionStack.Pop();
                                }
                            }

                            var pi = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos++);
                            if (EnumUtil.HasFlag(pi, FunctionParameterInformation.Condition))
                            {
                                var v = s.Pop().Compile();
                                PushResult(depChain._parsingContext, f, v);
                                fexp._latestConditionValue = GetCondition(v);
                                f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                            }
                            else if (fexp._latestConditionValue == ExpressionCondition.True || fexp._latestConditionValue == ExpressionCondition.False)
                            {
                                pi = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos);
                                if ((pi == FunctionParameterInformation.UseIfConditionIsFalse && fexp._latestConditionValue == ExpressionCondition.True)
                                   ||
                                   (pi == FunctionParameterInformation.UseIfConditionIsTrue && fexp._latestConditionValue == ExpressionCondition.False))
                                {
                                    f._tokenIndex = GetNextTokenPosFromCondition(f, fexp);
                                }
                            }
                            else if (fexp._latestConditionValue == ExpressionCondition.Error)
                            {
                                f._expressionStack.Push(Expression.Empty);
                                f._expressionStack.Push(Expression.Empty);
                                f._tokenIndex = fexp._endPos - 1;
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
                                if(funcExp.IsLet)
                                {
                                    f.IgnoreCaching = true;
                                }

                                if (PreExecFunc(depChain, f, funcExp) && returnAddresses)
                                {
                                    f._currentFunction = funcExp;
                                    f._tokenIndex--; //We should stay on this token when we continue on this formula.
                                    var a = funcExp._dependencyAddresses.ToArray();
                                    funcExp._dependencyAddresses.Clear();
                                    return a;
                                }
                            }
                            else
                            {
                                funcExp = f._currentFunction;
                                if (funcExp._dependencyAddresses.Count > 0 && returnAddresses)
                                {
                                    f._tokenIndex--; //We should stay on this token when we continue on this formula.
                                    var a = funcExp._dependencyAddresses.ToArray();
                                    funcExp._dependencyAddresses.Clear();
                                    return a;
                                }
                                f._currentFunction = null;
                            }

                            var r = ExecFunc(depChain, f, funcExp);
                            if(funcExp.ExecutesLambda)
                            {
                                // these functions will always invoke at least one LAMBDA
                                f.OnLambdaInvoked();
                            }
                            funcExp.OnDispose();
                            if (r.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic)
                            {
                                f._flags |= FormulaFlags.IsAlwaysDynamic;
                            }
                            if (funcExp.IsLambda && funcExp is not LambdaNameFunctionExpression && r.Result is LambdaCalculator clc)
                            {
                                f.LambdaSettings.LambdaStackNumbers.Push(s.Count);
                                f.LambdaSettings.NumberOfLambdaVariables.Push(clc.NumberOfVariables);
                            }
                            if (r.Address != null && returnAddresses)
                            {
                                if ((f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false) && r.Address != null)
                                {
                                    return GetCriteriaRange(depChain._parsingContext, f, r.Address.Clone());
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
                        if(fe is VariableFunctionExpression varFuncExp && varFuncExp.VariableScope != null)
                        {
                            depChain._parsingContext.VariableStorage.Push(varFuncExp.VariableScope);
                        }
                        if (fe._function == null)  //Function does not exists. Push #NAME?
                        {
                            f._tokenIndex = fe._endPos;
                            f._expressionStack.Push(new ErrorExpression(new CompileResult(eErrorType.Name), depChain._parsingContext));
                            break;
                        }
                        f._funcStack.Push(fe);
                        break;
                    case TokenType.Operator:
                        ApplyOperator(depChain._parsingContext, t, f);

                        if (returnAddresses && s.Count > 0 && s.Peek().Status == ExpressionStatus.IsAddress && (f._funcStack.Count == 0 || ShouldIgnoreAddress(f._funcStack.Peek()) == false))
                        {
                            var cr = s.Peek().Compile();
                            if (cr.Address != null)
                            {
                                return GetCriteriaRange(depChain._parsingContext, f, cr.Address);
                            }
                        }

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
                    case TokenType.Div0Error:
                        s.Push(ErrorExpression.Div0Error);
                        break;
                    case TokenType.GettingDataError:
                        s.Push(ErrorExpression.GettingDataError);
                        break;
                    case TokenType.Null:
                        s.Push(ErrorExpression.NullError);
                        break;
                    case TokenType.OpeningParenthesis:
                        f.OpenParenthesis();
                        break;
                    case TokenType.ClosingParenthesis:
                        f.CloseParenthesis(out bool shouldInvokeLambda);
                        if(shouldInvokeLambda)
                        {
                            var cr = LambdaInvoker.InvokeLambdaFunction(depChain, f);
                            if(f.LambdaSettings.NumberOfLambdaVariables.Count > 0)
                            {
                                f.LambdaSettings.NumberOfLambdaVariables.Pop();
                            }
                            if (f.LambdaSettings.LambdaArgsAdded.Count > 0)
                            {
                                f.LambdaSettings.LambdaArgsAdded.Pop();
                            }
                            if (cr != null)
                            {
                                PushResult(depChain._parsingContext, f, cr);
                            }
                        }
                        break;
                }

                f._tokenIndex++;
                if (f._tokenIndex == f._tokens.Count && returnAddresses)
                {
                    if (s.Count > 0 && s.Peek().Status == ExpressionStatus.IsAddress)
                    {
                        var cr = s.Peek().Compile();
                        addresses = [cr.Address]; //TODO:Check multi add
                    }
                }
            }
            return null;
        }

        private static FormulaRangeAddress[] GetCriteriaRange(ParsingContext ctx,RpnFormula f, FormulaRangeAddress address)
        {

            if (address.ExternalReferenceIx <=0 && f._funcStack.Count > 0)
            {
                var lfe = f._funcStack.Peek();
                var pi = lfe._function.ParametersInfo.GetParameterInfo(lfe._argPos);
                if (pi == FunctionParameterInformation.AdjustCriteriaParameterAddress)
                {
                    var q = new Queue<FormulaRangeAddress>();
                    lfe._function.GetNewParameterAddress(CreateArgumentsForParameterAddress(f, lfe),lfe._argPos, ctx, ref q);
                    return q.ToArray();
                }
            }
            return [address];
        }

        private static IList<CompileResult> CreateArgumentsForParameterAddress(RpnFormula f, FunctionExpression fe)
        {
            var ix = 0;
            var l = new List<CompileResult>();
            foreach(var e in f.ExpressionStack.Reverse())
            {
                if (fe._function.ParametersInfo.GetParameterInfo(ix)!=FunctionParameterInformation.AdjustParameterAddress)
                {
                    l.Add(e.Compile());
                }
                else
                {
                    l.Add(null);
                }
                ix++;
            }
            return l;
        }

        private static ExpressionCondition GetCondition(CompileResult v)
        {
            if (v.ResultValue is IRangeInfo ri)
            {
                var ret = ExpressionCondition.None;
                for (int r = 0; r < ri.Size.NumberOfRows; r++)
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
                            if (ret == ExpressionCondition.None)
                            {
                                ret = ExpressionCondition.Error;
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
            var v = e.Compile().ResultNumeric;
            v /= 100;
            f._expressionStack.Push(new DecimalExpression(new CompileResult(v, DataType.Decimal), context));
        }

        private static bool ShouldIgnoreAddress(FunctionExpression fe)
        {
            if (fe._function.ParametersInfo.HasNormalArguments == false)
            {
                var pi = fe._function.ParametersInfo.GetParameterInfo(fe._argPos);
                return (pi & (FunctionParameterInformation.IgnoreAddress | FunctionParameterInformation.AdjustParameterAddress)) != 0;
            }
            return false;
        }

        private static int GetNextTokenPosFromCondition(RpnFormula f, FunctionExpression fexp)
        {
            if (fexp._argPos < fexp.NumberOfArguments)
            {
                var fe = fexp._function.ParametersInfo.GetParameterInfo(fexp._argPos);
                while (fexp._argPos < fexp.NumberOfArguments && (
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
                if (fexp._argPos < fexp.NumberOfArguments)
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
            if(c2.Result is LambdaCalculator lc)
            {
                lc.BeginCalculation();
                var args = new List<Expression>();
                while(f._expressionStack.Count > 0)
                {
                    args.Add(f._expressionStack.Pop());
                }
                if(args.Any())
                {
                    var cr = args.First().Compile();
                    lc.SetVariableValue(0, cr.Result, cr.DataType, context);
                }
                c2 = lc.Execute(context);
                if(c2.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic)
                {
                    f._flags |= FormulaFlags.IsAlwaysDynamic;
                }
            }
            else if(c1.Result is LambdaCalculator lc2)
            {
                c1 = lc2.Execute(context);
                if (c1.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic)
                {
                    f._flags |= FormulaFlags.IsAlwaysDynamic;
                }
            }

            if (OperatorsDict.AllOperators.TryGetValue(opToken.Value, out IOperator op))
            {
                var result = op.Apply(c2, c1, context, true);
                if (result.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic)
                {
                    f._flags |= FormulaFlags.IsAlwaysDynamic;
                }
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
                    funcExp.Status = ExpressionStatus.CanCompile;
                    return funcExp.SetArguments(args, depChain._parsingContext);
                }
                else
                {
                    //Remove all function arguments from the stack
                    for (int i = 0; i < funcExp.NumberOfArguments && f._expressionStack.Count > 0; i++)
                    {
                        var si = f._expressionStack.Pop();
                    }
                    funcExp.Status = ExpressionStatus.IsCached;
                }
            }
            else
            {
                args = CompileFunctionArguments(f, funcExp);
                return funcExp.SetArguments(args, depChain._parsingContext);
            }
            return false;
        }

        private static CompileResult ExecFunc(RpnOptimizedDependencyChain depChain, RpnFormula f, FunctionExpression funcExp)
        {
            CompileResult result;
            funcExp.SetRpnFormula(f);
            if (funcExp.Status == ExpressionStatus.IsCached && !f.IgnoreCaching)
            {
                result = funcExp._cachedCompileResult;
            }
            else
            {
                result = funcExp.Compile();
                if (_cacheExpressions && !f.IgnoreCaching)
                {
                    funcExp._cachedCompileResult = result;
                    var key = funcExp.GetExpressionKey(f);
                    if (key != null)
                    {
                        funcExp.Status = ExpressionStatus.IsCached;
                        var cache = depChain.GetCache(f._ws);
                        cache.Add(key, result);
                    }
                }
            }
            if (funcExp._function != null && funcExp._function.ReturnsReference && result.Address != null && result.Address.FromRow > 0)
            {
                if (result.Result is InMemoryRange) //the result is an output from a reference function. Use the result instead.
                {
                    f._expressionStack.Push(new RangeExpression(result, depChain._parsingContext));
                }
                else
                {
                    f._expressionStack.Push(new RangeExpression(result.Address));
                }
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
                case DataType.WebImage:
                    f._expressionStack.Push(new WebImageExpression(result, context));
                    break;
                case DataType.LocalImage:   //References to local images
                    f._expressionStack.Push(new RangeExpression(result.Address));
                    break;
                case DataType.LambdaCalculation:
                    f._expressionStack.Push(new LambdaCalculationExpression(result, context));
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
            if (func.IsLambda)
            {
                for (int i = 0; i < func.NumberOfArguments; i++)
                {
                    var exp = s.Pop();
                    exp.Status |= ExpressionStatus.IsLambdaVariableDeclaration;
                    list.Insert(0, exp.Compile());
                }
            }
            else
            {
                var nArgs = func.IsLet ? ((LetFunctionExpression)func).NumberOfVariables + 1 : func.NumberOfArguments;
                for (int i = 0; i < nArgs && s.Count > 0; i++)
                {
                    var si = s.Pop();
                    if (si.ExpressionType != ExpressionType.Empty)
                    {
                        si.Status |= ExpressionStatus.FunctionArgument;
                    }
                    if(si is FunctionExpression fe1)
                    {
                        fe1.SetRpnFormula(f);
                    }
                    var cr = si.Compile();
                    list.Insert(0, cr);
                }
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