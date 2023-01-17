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
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class RpnNamedValueExpression : RpnExpression
    {
        short _externalReferenceIx;
        int _worksheetIx;
        internal INameInfo _name;
        bool _negate = false;
        public RpnNamedValueExpression(string name, ParsingContext parsingContext, short externalReferenceIx, int worksheetIx) : base(parsingContext)
        {
            _externalReferenceIx = externalReferenceIx;
            _name = Context.ExcelDataProvider.GetName(_externalReferenceIx, worksheetIx, name);
        }

        internal override ExpressionType ExpressionType => ExpressionType.NameValue;
        public override CompileResult Compile()
        {
            //var c = Context.Scopes.Current;
            
            //var cache = Context.AddressCache;
            //var cacheId = cache.GetNewId();
            if (_name == null) return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);

            if (_name.Value == null)
            {
                // check if there is a table with the name
                var table = Context.ExcelDataProvider.GetExcelTable(_name.Name);
                if(table != null)
                {
                    var ri = new RangeInfo(table.WorkSheet, table.Address);
                    return new AddressCompileResult(ri, DataType.ExcelRange, ri.Address);
                }

                return new CompileResult(eErrorType.Name);
            }

            if (_name.Value==null)
            {
                return new CompileResult(null, DataType.Empty);
            }

            if (_name.Value is IRangeInfo)
            {
                var range = (IRangeInfo)_name.Value;
                if (range.GetNCells()>1)
                {
                    return new AddressCompileResult(_name.Value, DataType.ExcelRange, range.Address);
                }
                else
                {                    
                    if (range.IsEmpty)
                    {
                        return new AddressCompileResult(null, DataType.Empty, range.Address);
                    }
                    var v = range.GetOffset(0,0);
                    return CompileResultFactory.Create(v, range.Address);
                }
            }
            else
            {                
                return CompileResultFactory.Create(_name.Value);
            }
            
            //return new CompileResultFactory().Create(result);
        }
        public override void Negate()
        {
            _negate = !_negate;
        }
        private ExcelExternalDefinedName GetExternalName()
        {
            ExcelWorkbook wb = Context.Package.Workbook;
            if (_externalReferenceIx >= 0 && _externalReferenceIx < wb.ExternalLinks.Count && wb.ExternalLinks[_externalReferenceIx].ExternalLinkType == ExternalReferences.eExternalLinkType.ExternalWorkbook)
            {
                var er = (ExcelExternalWorkbook)wb.ExternalLinks[_externalReferenceIx];
                if (_worksheetIx < 0)
                {
                    return er.CachedNames[_name.Name];
                }
                else
                {
                    return er.CachedWorksheets[_worksheetIx].CachedNames[_name.Name];
                }
            }
            return null;
        }
        public override FormulaRangeAddress GetAddress()
        {
            if(_name?.Value is IRangeInfo ri) 
            {
                return ri.Address;
            }
            return null;
        }
        internal override RpnExpressionStatus Status
        {
            get;
            set;
        } = RpnExpressionStatus.CanCompile;
    }
}
