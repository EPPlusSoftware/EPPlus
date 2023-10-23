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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Configuration of a <see cref="FormulaParser"/>
    /// </summary>
    public class ParsingConfiguration
    {
        /// <summary>
        /// Configures the formula calc engine to allow circular references.
        /// </summary>
        public bool AllowCircularReferences { get; internal set; }
        /// <summary>
        /// If EPPlus will should cache expressions or not. Default is true.
        /// </summary>
        public bool CacheExpressions { get; internal set; }
        /// <summary>
        /// In some functions EPPlus will round double values to 15 significant figures before the value is handled. This is an option for Excel compatibility.
        /// </summary>
        public PrecisionAndRoundingStrategy PrecisionAndRoundingStrategy { get; internal set; }

        /// <summary>
        /// The <see cref="IFormulaParserLogger"/> of the parser
        /// </summary>
        public IFormulaParserLogger Logger { get; private set; }

        ///// <summary>
        ///// The <see cref="IExpressionCompiler"/> of the parser
        ///// </summary>
        //public IExpressionCompiler ExpressionCompiler{ get; private set; }

        /// <summary>
        /// The <see cref="FunctionRepository"/> of the parser
        /// </summary>
        public FunctionRepository FunctionRepository{ get; private set; }

        /// <summary>
        /// Constructor
        /// </summary>
        private ParsingConfiguration() 
        {
            FunctionRepository = FunctionRepository.Create();
        }

        /// <summary>
        /// Factory method that creates an instance of this class
        /// </summary>
        /// <returns></returns>
        internal static ParsingConfiguration Create()
        {
            return new ParsingConfiguration();
        }

        /// <summary>
        /// Attaches a logger, errors and log entries will be written to the logger during the parsing process.
        /// </summary>
        /// <param name="logger"></param>
        /// <returns></returns>
        public ParsingConfiguration AttachLogger(IFormulaParserLogger logger)
        {
            Require.That(logger).Named("logger").IsNotNull();
            Logger = logger;
            return this;
        }

        /// <summary>
        /// if a logger is attached it will be removed.
        /// </summary>
        /// <returns></returns>
        public ParsingConfiguration DetachLogger()
        {
            Logger?.Dispose();
            Logger = null;
            return this;
        }
    }
}
