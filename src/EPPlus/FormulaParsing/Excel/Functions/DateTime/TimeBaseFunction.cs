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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    internal abstract class TimeBaseFunction : ExcelFunction
    {
        public TimeBaseFunction()
        {
            TimeStringParser = new TimeStringParser();
        }

        protected TimeStringParser TimeStringParser
        {
            get;
            private set;
        }

        protected double SerialNumber
        {
            get;
            private set;
        }

        public void ValidateAndInitSerialNumber(IEnumerable<FunctionArgument> arguments)
        {
            SerialNumber = (double)ArgToDecimal(arguments, 0);
        }

        protected double SecondsInADay
        {
            get{ return 24 * 60 * 60; }
        }

        protected double GetTimeSerialNumber(double seconds)
        {
            return seconds / SecondsInADay;
        }

        protected double GetSeconds(double serialNumber)
        {
            return serialNumber * SecondsInADay;
        }

        protected double GetHour(double serialNumber)
        {
            var seconds = GetSeconds(serialNumber);
            return (int)seconds / (60 * 60);
        }

        protected double GetMinute(double serialNumber)
        {
            var seconds = GetSeconds(serialNumber);
            seconds -= GetHour(serialNumber) * 60 * 60;
            return (seconds - (seconds % 60)) / 60;
        }

        protected double GetSecond(double serialNumber)
        {
            return GetSeconds(serialNumber) % 60;
        }
    }
}
