using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.Core.RichValues
{
    internal class RichValueErrorManager
    {
        public RichValueErrorManager(ExcelPackage package, ExcelWorksheet ws)
        {
            _package = package;
            _ws = ws;
            _richDataDb = _package.Workbook.RichData.Db;
            _richDataStore = new RichDataStore(ws);
            _metadataStore = _ws._metadataStore;
        }

        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _ws;
        private readonly RichDataDatabase _richDataDb;
        private readonly RichDataStore _richDataStore;
        private readonly CellStore<MetaDataReference> _metadataStore;

        internal object GetErrorFromMetaData(int row, int col, object v)
        {
            var md = _metadataStore.GetValue(row, col);
            if (md.vm > 0)
            {
                v = GetErrorFromMetaData(md, v);
            }
            return v;
        }

        //
        internal object GetErrorFromMetaData(MetaDataReference md, object v)
        {
            var rdValue = _richDataStore.GetRichValue(md.vm);
            var error = RichValueErrorFactory.CreateRichValueErrorFromRichData(rdValue, _package.Workbook.IndexStore, _package.Workbook.RichData.Db);
            if (error != null)
            {
                switch(error.ErrorType)
                {
                    case 4:
                        return ErrorValues.NameError;
                    case 8:
                        var spillError = error.As.ErrorSpill;
                        if (spillError != null && spillError.RwOffset > -1 && spillError.ColOffset > 0)
                        {
                            return new ExcelRichDataErrorValue(spillError.RwOffset ?? 0, spillError.ColOffset ?? 0);
                        }
                        return new ExcelRichDataErrorValue(0, 0);

                    case 13:
                        return ErrorValues.CalcError;
                    default:  //We can implement other error types here later, See MS-XLSX 2.3.6.1.3
                        return v;
                }
            }
            return v;
        }

        internal void SetMetaDataForError(CellStoreEnumerator<ExcelValue> cse, ExcelErrorValue error)
        {
            var metadata = _package.Workbook.Metadata;
            if(_richDataStore.HasRichData(cse.Row, cse.Column, out MetaDataReference md))
            {
                var richValue = _richDataStore.GetRichValue(cse.Row, cse.Column);
                if (richValue == null || IsMdSameError(richValue, error)) return;
            }
            var newRv = default(ExcelRichValue);
            switch (error.Type)
            {
                case eErrorType.Spill:
                    var spillError = (ExcelRichDataErrorValue)error;
                    if (spillError.IsPropagated)
                    {
                        newRv = CreatePropagated(eErrorType.Spill);
                    }
                    else
                    {
                       newRv = CreateErrorSpill(spillError);
                    }
                    break;
                case eErrorType.Calc:
                    newRv = CreateError(eErrorType.Calc, 1);
                    break;
                default:
                    return;
            }
            if(newRv != null)
            {
                _richDataStore.AddRichData(cse.Row, cse.Column, newRv);
            }
        }

        private bool IsMdSameError(ExcelRichValue richValue, ExcelErrorValue error)
        {
            if(richValue == null) return false;
            if (richValue.Structure.Type == StructureTypes.Error)
            {
                var rdErrorBase = RichValueErrorFactory.CreateRichValueErrorFromRichData(richValue, _package.Workbook.IndexStore, _package.Workbook.RichData.Db);
                switch (error.Type)
                {
                    case eErrorType.Calc:
                        return rdErrorBase.ErrorType == 13;
                    case eErrorType.Spill:
                        var rdError = (ExcelRichDataErrorValue)error;
                        var spillError = rdErrorBase.As.ErrorSpill;
                        if(spillError != null)
                        {
                            return spillError.AreEqual(8, rdError.SpillColOffset, rdError.SpillRowOffset);
                        }
                        break;
                    default:
                        return false;

                }
            }
            return false;
        }

        private ErrorPropagatedRichValue CreatePropagated(eErrorType errorType)
        {
            var item = new ErrorPropagatedRichValue(_richDataDb)
            {
                Propagated = "1"
            };
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.ErrorType = RichDataErrorType.Calc;
                    break;
                case eErrorType.Spill:
                    item.ErrorType = RichDataErrorType.Spill;
                    break;

            }
            return item;
        }

        internal ErrorWithSubTypeRichValue CreateError(eErrorType errorType, int subType)
        {
            var item = new ErrorWithSubTypeRichValue(_richDataDb)
            {
                SubType = subType
            };
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.ErrorType = RichDataErrorType.Calc;
                    break;
                case eErrorType.Spill:
                    item.ErrorType = RichDataErrorType.Spill;
                    break;

            }
            return item;
        }

        internal ErrorSpillRichValue CreateErrorSpill(ExcelRichDataErrorValue spillError)
        {
            var item = new ErrorSpillRichValue(_richDataDb)
            {
                ColOffset = spillError.SpillColOffset,
                RwOffset = spillError.SpillRowOffset,
                SubType = 1,
                ErrorType = RichDataErrorType.Spill
            };
            return item;
        }
    }
}
