using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
    internal class RpnOptimizedDependencyChain
    {
        internal List<RpnFormula> _formulas = new List<RpnFormula>();
        internal Stack<RpnFormula> _formulaStack=new Stack<RpnFormula>();
        internal Dictionary<int, RangeHashset> accessedRanges = new Dictionary<int, RangeHashset>();
        internal Dictionary<int, QuadTree<int>> formulaRangeReferences = new Dictionary<int, QuadTree<int>>();
        internal HashSet<ulong> processedCells = new HashSet<ulong>();
        internal List<CircularReference> _circularReferences = new List<CircularReference>();
        internal ISourceCodeTokenizer _tokenizer;
        internal FormulaExecutor _formulaExecutor;
        internal ParsingContext _parsingContext;
        internal List<int> _startOfChain = new List<int>();
        internal bool HasDynamicArrayFormula=false;
        internal Dictionary<int, Dictionary<string, CompileResult>> _expressionCache = new Dictionary<int, Dictionary<string, CompileResult>>();
        internal bool HasAnyArrayFormula { get; set; } = false;
        public RpnOptimizedDependencyChain(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            _tokenizer = SourceCodeTokenizer.Default;
            _parsingContext = wb.FormulaParser.ParsingContext;
            _formulaExecutor = new FormulaExecutor(_parsingContext);

            var parser = wb.FormulaParser;
            var filterInfo = new FilterInfo(wb);
            parser.InitNewCalc(filterInfo);

            wb.FormulaParser.Configure(config =>
            {
                config.AllowCircularReferences = options.AllowCircularReferences;
                config.CacheExpressions = options.CacheExpressions;
                config.PrecisionAndRoundingStrategy = options.PrecisionAndRoundingStrategy;
            });

        }

        internal void AddFormulaToChain(RpnFormula f)
        {
            QuadTree<int> qr;
            var ix = f._ws?.IndexInList ?? short.MaxValue;
            if (formulaRangeReferences.TryGetValue(ix, out qr) == false)
            {
                if (f._ws == null)
                {
                    qr = new QuadTree<int>(1,1, _parsingContext.Package.Workbook.Names.Count, 1);
                }
                else
                {
                    if(f._ws.Dimension==null)
                    {
                        qr = new QuadTree<int>(QuadRange.MinSize, QuadRange.MinSize, QuadRange.MinSize, QuadRange.MinSize);
                    }
                    else
                    {
                        qr = new QuadTree<int>(f._ws.Dimension);
                    }                    
                }
                formulaRangeReferences.Add(ix, qr);  
            }
            foreach(var e in f._expressions)
            {
                if((e.Value.Status & ExpressionStatus.IsAddress) == ExpressionStatus.IsAddress)
                {
                    foreach (var a in e.Value.GetAddress())
                    {
                        qr.Add(new QuadRange(a), _formulas.Count);
                    }
                }
            }
            _formulas.Add(f);
        }

        internal RpnOptimizedDependencyChain Execute()
        {
            return RpnFormulaExecution.Execute(_parsingContext.Package.Workbook, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws)
        {
            return RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
        }
        internal RpnOptimizedDependencyChain Execute(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            return RpnFormulaExecution.Execute(ws, options);
        }

        internal Dictionary<string, CompileResult> GetCache(ExcelWorksheet ws)
        {
            var ix = ws == null ? -1 : ws.IndexInList;

            if(!_expressionCache.TryGetValue(ix, out Dictionary<string, CompileResult> cache))
            {
                cache = new Dictionary<string, CompileResult>();
                _expressionCache.Add(ix, cache);
            }
            return cache;
        }

        //Adds the position where a chain of formulas start.
        internal void StartOfChain()
        {
            _startOfChain.Add(_formulas.Count);
        }
    }
}
