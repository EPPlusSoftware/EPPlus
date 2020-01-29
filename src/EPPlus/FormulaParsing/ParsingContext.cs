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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Logging;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Parsing context
    /// </summary>
    public class ParsingContext : IParsingLifetimeEventHandler
    {
        private ParsingContext() { }

        /// <summary>
        /// The <see cref="FormulaParser"/> of the current context.
        /// </summary>
        public FormulaParser Parser { get; set; }

        /// <summary>
        /// The <see cref="ExcelDataProvider"/> is an abstraction on top of
        /// Excel, in this case EPPlus.
        /// </summary>
        public ExcelDataProvider ExcelDataProvider { get; set; }

        /// <summary>
        /// Utility for handling addresses
        /// </summary>
        public RangeAddressFactory RangeAddressFactory { get; set; }

        /// <summary>
        /// <see cref="INameValueProvider"/> of the current context
        /// </summary>
        public INameValueProvider NameValueProvider { get; set; }

        /// <summary>
        /// Configuration
        /// </summary>
        public ParsingConfiguration Configuration { get; set; }

        /// <summary>
        /// Scopes, a scope represents the parsing of a cell or a value.
        /// </summary>
        public ParsingScopes Scopes { get; private set; }

        /// <summary>
        /// Address cache
        /// </summary>
        /// <seealso cref="ExcelAddressCache"/>
        public ExcelAddressCache AddressCache { get; private set; }

        /// <summary>
        /// Returns true if a <see cref="IFormulaParserLogger"/> is attached to the parser.
        /// </summary>
        public bool Debug
        {
            get { return Configuration.Logger != null; }
        }

        /// <summary>
        /// Factory method.
        /// </summary>
        /// <returns></returns>
        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            context.Scopes = new ParsingScopes(context);
            context.AddressCache = new ExcelAddressCache();
            return context;
        }

        void IParsingLifetimeEventHandler.ParsingCompleted()
        {
            AddressCache.Clear();
        }
    }
}
