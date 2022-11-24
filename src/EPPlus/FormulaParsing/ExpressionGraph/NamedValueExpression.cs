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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using OfficeOpenXml.ExternalReferences;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class NamedValueExpression : ExpressionWithParent
    {
        FormulaAddressBase _locationInfo;
        public NamedValueExpression(string expression, ParsingContext parsingContext, ref FormulaAddressBase locationInfo)
            : base(expression, parsingContext)
        {
            _parsingContext = parsingContext;
            _locationInfo = locationInfo??new FormulaAddressBase();
        }

        private readonly ParsingContext _parsingContext;

        internal override ExpressionType ExpressionType => ExpressionType.NameValue;
        public override bool IsGroupedExpression => false;
        public override CompileResult Compile()
        {
            var c = _parsingContext.Scopes.Current;
            var name = _parsingContext.ExcelDataProvider.GetName(_locationInfo.ExternalReferenceIx, _locationInfo.WorksheetIx, ExpressionString);
            
            var cache = _parsingContext.AddressCache;
            var cacheId = cache.GetNewId();
            
            if (name == null)
            {
                // check if there is a table with the name
                var table = _parsingContext.ExcelDataProvider.GetExcelTable(ExpressionString);
                if(table != null)
                {
                    var ri = new RangeInfo(table.WorkSheet, table.Address);
                    cache.Add(cacheId, ri.Address.ToString());
                    return new CompileResult(ri, DataType.Enumerable, cacheId);
                }

                return new CompileResult(eErrorType.Name);
            }

            if (name.Value==null)
            {
                return new CompileResult(null, DataType.Empty, cacheId);
            }

            if (name.Value is IRangeInfo)
            {
                var range = (IRangeInfo)name.Value;
                if (range.GetNCells()>1)
                {
                    return new AddressCompileResult(name.Value, DataType.Enumerable, range.Address);
                }
                else
                {                    
                    if (range.IsEmpty)
                    {
                        return new AddressCompileResult(null, DataType.Empty, range.Address);
                    }
                    return CompileResultFactory.Create(range.First().Value, cacheId, range.Address);
                }
            }
            else
            {                
                return CompileResultFactory.Create(name.Value, cacheId);
            }
            
            //return new CompileResultFactory().Create(result);
        }

        private ExcelExternalDefinedName GetExternalName()
        {
            ExcelWorkbook wb = _parsingContext.Package.Workbook;
            var erIx = _locationInfo.ExternalReferenceIx - 1;
            if (erIx >= 0 && erIx < wb.ExternalLinks.Count && wb.ExternalLinks[erIx].ExternalLinkType == ExternalReferences.eExternalLinkType.ExternalWorkbook)
            {
                var er = (ExcelExternalWorkbook)wb.ExternalLinks[erIx];                
                if (_locationInfo.WorksheetIx < 0)
                {
                    return er.CachedNames[ExpressionString];
                }
                else
                {
                    return er.CachedWorksheets[_locationInfo.WorksheetIx].CachedNames[ExpressionString];
                }
            }
            return null;
        }
        internal override Expression Clone()
        {
            return CloneMe(new NamedValueExpression(ExpressionString, Context, ref _locationInfo));
        }
    }
}
