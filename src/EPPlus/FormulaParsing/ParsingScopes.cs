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
    /// This class implements a stack on which instances of <see cref="ParsingScope"/>
    /// are put. Each ParsingScope represents the parsing of an address in the workbook.
    /// </summary>
    public class ParsingScopes
    {
        private readonly IParsingLifetimeEventHandler _lifetimeEventHandler;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="lifetimeEventHandler">An instance of a <see cref="IParsingLifetimeEventHandler"/></param>
        public ParsingScopes(IParsingLifetimeEventHandler lifetimeEventHandler)
        {
            _lifetimeEventHandler = lifetimeEventHandler;
        }
        private Stack<ParsingScope> _scopes = new Stack<ParsingScope>();

        /// <summary>
        /// Creates a new <see cref="ParsingScope"/> and puts it on top of the stack.
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public virtual ParsingScope NewScope(FormulaRangeAddress address)
        {
            ParsingScope scope;
            if (_scopes.Count() > 0)
            {
                scope = new ParsingScope(this, _scopes.Peek(), address);
            }
            else
            {
                scope = new ParsingScope(this, address);
            }
            _scopes.Push(scope);
            return scope;
        }


        /// <summary>
        /// The current parsing scope.
        /// </summary>
        public virtual ParsingScope Current
        {
            get { return _scopes.Count() > 0 ? _scopes.Peek() : null; }
        }

        /// <summary>
        /// Removes the current scope, setting the calling scope to current.
        /// </summary>
        /// <param name="parsingScope"></param>
        public virtual void KillScope(ParsingScope parsingScope)
        {
            _scopes.Pop();
            if (_scopes.Count() == 0)
            {
                _lifetimeEventHandler.ParsingCompleted();
            }
        }
    }
}
