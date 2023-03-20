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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    /// <summary>
    /// This class provides access to <see cref="SeparatorHandler"/>s - classes that exposes functionatlity
    /// needed when parsing strings to tokens.
    /// </summary>
    internal class TokenSeparatorHandler
    {
        public TokenSeparatorHandler(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider)
            : this(new SeparatorHandler[]
                {
                    new StringHandler(),
                    new BracketHandler(),
                    new SheetnameHandler(),
                    new MultipleCharSeparatorHandler(tokenSeparatorProvider, nameValueProvider),
                    new DefinedNameAddressHandler(nameValueProvider),
                    new ExponentialNumberHandler()
                }){}

        public TokenSeparatorHandler(params SeparatorHandler[] handlers)
        {
            _handlers = handlers;
        }

        private readonly SeparatorHandler[] _handlers;

        /// <summary>
        /// Handles a tokenseparator.
        /// </summary>
        /// <param name="c"></param>
        /// <param name="tokenSeparator"></param>
        /// <param name="context"></param>
        /// <param name="tokenIndexProvider"></param>
        /// <returns>Returns true if the tokenseparator was handled.</returns>
        public bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            foreach(var handler in _handlers)
            {
                if(handler.Handle(c, tokenSeparator, context, tokenIndexProvider))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
