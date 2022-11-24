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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Represents a parsing of a single input or workbook addrses.
    /// </summary>
    public class ParsingScope : IDisposable
    {
        private readonly ParsingScopes _parsingScopes;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="parsingScopes"></param>
        /// <param name="address"></param>
        public ParsingScope(ParsingScopes parsingScopes, FormulaRangeAddress address)
            : this(parsingScopes, null, address)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="parsingScopes"></param>
        /// <param name="parent"></param>
        /// <param name="address"></param>
        public ParsingScope(ParsingScopes parsingScopes, ParsingScope parent, FormulaRangeAddress address)
        {
            _parsingScopes = parsingScopes;
            Parent = parent;
            Address = address;
            ScopeId = Guid.NewGuid();
        }

        /// <summary>
        /// Id of the scope.
        /// </summary>
        public Guid ScopeId { get; private set; }

        /// <summary>
        /// The calling scope.
        /// </summary>
        public ParsingScope Parent { get; private set; }

        /// <summary>
        /// The address of the cell currently beeing parsed.
        /// </summary>
        public FormulaRangeAddress Address { get; private set; }

        /// <summary>
        /// True if the current scope is a Subtotal function being executed.
        /// </summary>
        public bool IsSubtotal { get; set; }

        /// <summary>
        /// Disposes this instance
        /// </summary>
        public void Dispose()
        {
            _parsingScopes.KillScope(this);
        }
    }
}
