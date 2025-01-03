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
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Address compile result
    /// </summary>
    public class AddressCompileResult : CompileResult
    {
        /// <summary>
        /// Address result
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        public AddressCompileResult(object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType)
        {
            Address = address;
        }
        /// <summary>
        /// Address result without address
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        public AddressCompileResult(object result, DataType dataType) : base(result, dataType)
        { 

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="error"></param>
        public AddressCompileResult(eErrorType error) : base(error)
        {

        }
        /// <summary>
        /// Address compile result
        /// </summary>
        /// <param name="errorValue"></param>
        public AddressCompileResult(ExcelErrorValue errorValue) : base(errorValue)
        {

        }
        /// <summary>
        /// Address
        /// </summary>
        public override FormulaRangeAddress Address
        {
            get;
        }
        /// <summary>
        /// ResultType
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                if(Address==null)
                {
                    return base.ResultType;
                }
                else if(ResultValue != null && ResultValue is ExcelCellPicture ecp)
                {
                    if(ecp.PictureType == ExcelCellPictureTypes.LocalImage)
                    {
                        return CompileResultType.LocalImage;
                    }
                    else
                    {
                        return CompileResultType.WebImage;
                    }
                }
                return CompileResultType.RangeAddress;
            }
        }
    }
}
