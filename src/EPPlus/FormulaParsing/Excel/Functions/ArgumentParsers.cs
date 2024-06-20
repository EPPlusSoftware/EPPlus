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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Argument parsers
    /// </summary>
    internal class ArgumentParsers
    {
        private static object _syncRoot = new object();
        private readonly Dictionary<DataType, ArgumentParser> _parsers = new Dictionary<DataType, ArgumentParser>();
        private readonly ArgumentParserFactory _parserFactory;

        /// <summary>
        /// Empty constructor
        /// </summary>
        public ArgumentParsers()
            : this(new ArgumentParserFactory())
        {

        }
        /// <summary>
        /// Factory constructor.
        /// </summary>
        /// <param name="factory">Factory cannot be null</param>
        public ArgumentParsers(ArgumentParserFactory factory)
        {
            Require.That(factory).Named("argumentParserfactory").IsNotNull();
            _parserFactory = factory;
        }
        /// <summary>
        /// Get parser of type
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        public ArgumentParser GetParser(DataType dataType)
        {
            if (!_parsers.ContainsKey(dataType))
            {
                lock (_syncRoot)
                {
                    if (!_parsers.ContainsKey(dataType))
                    {
                        _parsers.Add(dataType, _parserFactory.CreateArgumentParser(dataType));
                    }
                }
            }
            return _parsers[dataType];
        }
    }
}
