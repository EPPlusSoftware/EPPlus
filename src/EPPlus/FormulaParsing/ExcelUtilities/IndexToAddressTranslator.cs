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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    internal class IndexToAddressTranslator
    {
        internal IndexToAddressTranslator(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ExcelReferenceType.AbsoluteRowAndColumn)
        {

        }

        internal IndexToAddressTranslator(ExcelDataProvider excelDataProvider, ExcelReferenceType referenceType)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _excelReferenceType = referenceType;
        }

        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ExcelReferenceType _excelReferenceType;

        protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
        {

            if (iColumnNumber < 1)
            {
                //throw new Exception("Column number is out of range");
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return fixedCol ? "$" + sCol : sCol;
        }

        public string ToAddress(int col, int row)
        {
            var fixedCol = _excelReferenceType == ExcelReferenceType.AbsoluteRowAndColumn ||
                           _excelReferenceType == ExcelReferenceType.RelativeRowAbsoluteColumn;
            var colString = GetColumnLetter(col, fixedCol);
            return colString + GetRowNumber(row);
        }

        private string GetRowNumber(int rowNo)
        {
            var retVal = rowNo < (_excelDataProvider.ExcelMaxRows) ? rowNo.ToString() : string.Empty;
            if (!string.IsNullOrEmpty(retVal))
            {
                switch (_excelReferenceType)
                {
                    case ExcelReferenceType.AbsoluteRowAndColumn:
                    case ExcelReferenceType.AbsoluteRowRelativeColumn:
                        return "$" + retVal;
                    default:
                        return retVal;
                }
            }
            return retVal;
        }
    }
}
