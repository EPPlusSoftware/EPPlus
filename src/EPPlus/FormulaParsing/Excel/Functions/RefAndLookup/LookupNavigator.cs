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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class LookupNavigator
    {
        protected readonly LookupDirection Direction;
        protected readonly LookupArguments Arguments;
        protected readonly ParsingContext ParsingContext;



        public LookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            Require.That(parsingContext).Named("parsingContext").IsNotNull();
            Require.That(parsingContext.ExcelDataProvider).Named("parsingContext.ExcelDataProvider").IsNotNull();
            Direction = direction;
            Arguments = arguments;
            ParsingContext = parsingContext;
        }

        public abstract int Index
        {
            get;
        }

        public abstract bool MoveNext();

        public abstract object CurrentValue
        {
            get;
        }

        public abstract object GetLookupValue();
    }
}
