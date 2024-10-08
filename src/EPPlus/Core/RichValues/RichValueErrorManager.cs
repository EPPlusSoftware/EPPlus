using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Errors;
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
            _richData = _package.Workbook.RichData;
            _richDataStore = new RichDataStore(ws);
            _metadataStore = _ws._metadataStore;
        }

        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _ws;
        private readonly ExcelRichData _richData;
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

            var error = rdValue.As.Type<ErrorRichValueBase>();
            if(error != null)
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

            //var metaData = _package.Workbook.Metadata;
            //var valueMetaData = metaData.ValueMetadata[md.vm - 1];
            //var valueRecord = valueMetaData.Records[0];
            //var type = _richDataStore.GetMetadataType(md.vm);
            //if (type.Name.Equals("XLRICHVALUE"))
            //{
            //    var fmd = metaData.FutureMetadata[type.Name];
            //    var ix = fmd.Types[valueRecord.ValueIndex].AsRichData.Index;

            //    var rdValue = _richData.Values.Items[ix];

            //    var errorTypeIndex = rdValue.Structure.Keys.FindIndex(x => x.Name.Equals("errorType"));
            //    if (errorTypeIndex >= 0)
            //    {
            //        switch (int.Parse(rdValue.Values[errorTypeIndex]))
            //        {
            //            case 4:
            //                return ErrorValues.NameError;
            //            case 8:
            //                var rowOffsetIndex = rdValue.Structure.Keys.FindIndex(x => x.Name.Equals("rwOffset"));
            //                var colOffsetIndex = rdValue.Structure.Keys.FindIndex(x => x.Name.Equals("colOffset"));
            //                if (rowOffsetIndex > -1 && colOffsetIndex > 0)
            //                {
            //                    return new ExcelRichDataErrorValue(int.Parse(rdValue.Values[rowOffsetIndex]), int.Parse(rdValue.Values[colOffsetIndex]));
            //                }
            //                else
            //                {
            //                    return new ExcelRichDataErrorValue(0, 0);
            //                }
            //            case 13:
            //                return ErrorValues.CalcError;
            //            default:    //We can implement other error types here later, See MS-XLSX 2.3.6.1.3
            //                return v;

            //        }
            //    }
            //}
            //return v;
        }

        internal void SetMetaDataForError(CellStoreEnumerator<ExcelValue> cse, ExcelErrorValue error)
        {
            var metadata = _package.Workbook.Metadata;
            //var md = _ws._metadataStore.GetValue(cse.Row, cse.Column);
            if(_richDataStore.HasRichData(cse.Row, cse.Column, out MetaDataReference md))
            {
                var richValue = _richDataStore.GetRichValue(cse.Row, cse.Column);
                if (richValue == null || IsMdSameError(richValue, error)) return;
            }

            //if (md.vm >= 0 && IsMdSameError(metadata, md, error, cse.Row, cse.Column))
            //{
            //    return;
            //}
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
                //md.vm = vm;
                //_ws._metadataStore.SetValue(cse.Row, cse.Column, md);
            }
            //metadata.CreateRichValueMetadata(_richData, out int newVm);
            //md.vm = newVm;
            //_ws._metadataStore.SetValue(cse.Row, cse.Column, md);
        }

        //private bool IsMdSameError(ExcelMetadata metadata, MetaDataReference md, ExcelErrorValue error, int row, int column)
        private bool IsMdSameError(ExcelRichValue richValue, ExcelErrorValue error)
        {
            //if (md.vm == 0 || md.vm >= metadata.ValueMetadata.Count) return false;

            //var richData = _richDataStore.GetRichValue(md.vm);
            if(richValue == null) return false;
            if (richValue.Structure.Type == StructureTypes.Error)
            {
                var rdErrorBase = richValue.As.Type<ErrorRichValueBase>();
                switch (error.Type)
                {
                    case eErrorType.Calc:
                        return rdErrorBase.ErrorType == 13;
                    case eErrorType.Spill:
                        var rdError = (ExcelRichDataErrorValue)error;
                        var spillError = richValue.As.ErrorSpill;
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

            //var vm = metadata.ValueMetadata[md.vm - 1];

            //if (vm.Records.Count > 0 && vm.Records[0].ValueIndex >= 0)
            //{
            //    if (_richData.Values.Items.Count > vm.Records[0].ValueIndex)
            //    {
            //        var rd = _richData.Values.Items[vm.Records[0].ValueIndex];
            //        if (rd.Structure.Type.Equals(StructureTypes.Error))
            //        {
            //            ;
            //            switch (error.Type)
            //            {
            //                case eErrorType.Calc:
            //                    if (rd.Values[0] == "13")
            //                    {
            //                        return true;
            //                    }
            //                    break;
            //                case eErrorType.Spill:
            //                    var rdError = (ExcelRichDataErrorValue)error;
            //                    if (rd.HasValue(["errorType", "colOffset", "rwOffset"], ["8", rdError.SpillColOffset.ToString(CultureInfo.InvariantCulture), rdError.SpillColOffset.ToString(CultureInfo.InvariantCulture)]))
            //                    {
            //                        return true;
            //                    }
            //                    break;
            //            }
            //        }
            //    }
            //}
            //return false;
        }

        private ErrorPropagatedRichValue CreatePropagated(eErrorType errorType)
        {
            var item = new ErrorPropagatedRichValue(_package.Workbook)
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
            //_richData.Values.Items.Add(item);
            return item;
        }

        internal ErrorWithSubTypeRichValue CreateError(eErrorType errorType, int subType)
        {
            var item = new ErrorWithSubTypeRichValue(_package.Workbook)
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
            //_richData.Values.Items.Add(item);
            return item;
        }

        internal ErrorSpillRichValue CreateErrorSpill(ExcelRichDataErrorValue spillError)
        {
            var item = new ErrorSpillRichValue(_package.Workbook)
            {
                ColOffset = spillError.SpillColOffset,
                RwOffset = spillError.SpillRowOffset,
                SubType = 1,
                ErrorType = RichDataErrorType.Spill
            };
            //_richData.Values.Items.Add(item);
            return item;
        }
    }
}
