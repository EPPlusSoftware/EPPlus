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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Handles translations from Spreadsheet addresses to 0-based numeric index.
    /// </summary>
    public class AddressTranslator
    {
        public enum RangeCalculationBehaviour
        {
            FirstPart,
            LastPart
        }

        private readonly ExcelDataProvider _excelDataProvider;

        public AddressTranslator(ExcelDataProvider excelDataProvider)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
        }

        /// <summary>
        /// Translates an address in format "A1" to col- and rowindex.
        /// 
        /// If the supplied address is a range, the address of the first part will be calculated.
        /// </summary>
        /// <param name="address"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        public virtual void ToColAndRow(string address, out int col, out int row)
        {
            ToColAndRow(address, out col, out row, RangeCalculationBehaviour.FirstPart);
        }

        /// <summary>
        /// Translates an address in format "A1" to col- and rowindex.
        /// </summary>
        /// <param name="address"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <param name="behaviour"></param>
        public virtual void ToColAndRow(string address, out int col, out int row, RangeCalculationBehaviour behaviour)
        {
            address = Utils.ConvertUtil._invariantTextInfo.ToUpper(address);
            var alphaPart = GetAlphaPart(address);
            col = 0;
            var nLettersInAlphabet = 26;
            for (int x = 0; x < alphaPart.Length; x++)
            {
                var pos = alphaPart.Length - x - 1;
                var currentNumericValue = GetNumericAlphaValue(alphaPart[x]);
                col += (nLettersInAlphabet * pos * currentNumericValue);
                if (pos == 0)
                {
                    col += currentNumericValue;
                }
            }
            //col--;
            //row = GetIntPart(address) - 1 ?? GetRowIndexByBehaviour(behaviour);
            row = GetIntPart(address) ?? GetRowIndexByBehaviour(behaviour);

        }

        private int GetRowIndexByBehaviour(RangeCalculationBehaviour behaviour)
        {
            if (behaviour == RangeCalculationBehaviour.FirstPart)
            {
                return 1;
            }
            return _excelDataProvider.ExcelMaxRows;
        }

        private int GetNumericAlphaValue(char c)
        {
            return (int)c - 64;
        }

        private string GetAlphaPart(string address)
        {
            return Regex.Match(address, "[A-Z]+").Value;
        }

        private int? GetIntPart(string address)
        {
            if (Regex.IsMatch(address, "[0-9]+"))
            {
                return int.Parse(Regex.Match(address, "[0-9]+").Value);
            }
            return null;
        }
    }
}
