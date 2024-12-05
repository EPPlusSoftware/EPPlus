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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Logging;
using NvProvider = OfficeOpenXml.FormulaParsing.NameValueProvider;
using System;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Excel;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Parsing context
    /// </summary>
    public class ParsingContext 
    {
        private ParsingContext(ExcelPackage package) {
            SubtotalAddresses = new HashSet<ulong>();
            Package = package;
        }

        internal FunctionCompilerFactory FunctionCompilerFactory
        {
            get; private set;
        }

        /// <summary>
        /// The <see cref="FormulaParser"/> of the current context.
        /// </summary>
        public FormulaParser Parser { get; set; }

        /// <summary>
        /// The <see cref="ExcelDataProvider"/> is an abstraction on top of
        /// Excel, in this case EPPlus.
        /// </summary>
        internal ExcelDataProvider ExcelDataProvider { get; set; }

        /// <summary>
        /// The <see cref="ExcelPackage"/> where the calculation is done.
        /// </summary>
        internal ExcelPackage Package { get; private set; }

        /// <summary>
        /// Indicates how hidden cells should be handled by the currently
        /// executing function.
        /// </summary>
        internal HiddenCellHandlingCategory HiddenCellBehaviour
        {
            get; set;
        }

        /// <summary>
        /// Utility for handling addresses
        /// </summary>
        internal RangeAddressFactory RangeAddressFactory { get; set; }

        /// <summary>
        /// <see cref="INameValueProvider"/> of the current context
        /// </summary>
        public INameValueProvider NameValueProvider { get; set; }

        /// <summary>
        /// Configuration
        /// </summary>
        public ParsingConfiguration Configuration { get; set; }
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
        /// <param name="package">The ExcelPackage where calculation is done</param>
        /// <returns></returns>
        public static ParsingContext Create(ExcelPackage package)
        {
            var context = new ParsingContext(package);
            context.Configuration = ParsingConfiguration.Create();
            //context.Scopes = new ParsingScopes(context);
            //context.AddressCache = new ExcelAddressCache();
            context.NameValueProvider = NvProvider.Empty;
            context.FunctionCompilerFactory = new FunctionCompilerFactory(context.Configuration.FunctionRepository);
            return context;
        }

        /// <summary>
        /// Factory method
        /// </summary>
        /// <returns></returns>
        public static ParsingContext Create()
        {
            return Create(null);
        }

        internal int GetWorksheetIndex(string wsName)
        {
            if(string.IsNullOrEmpty(wsName))
            {
                return CurrentCell.WorksheetIx;
            }
            else
            {
                return Package.Workbook.Worksheets.GetPositionByToken(wsName);
            }
        }

        internal ExcelExternalWorkbook GetExternalWoorkbook(int externalReferenceIx)
        {
            return Package.Workbook.ExternalLinks[externalReferenceIx - 1] as ExcelExternalWorkbook;
        }

        internal HashSet<ulong> SubtotalAddresses { get; private set; }
        /// <summary>
        /// Represents the current cell 
        /// </summary>
        public FormulaCellAddress CurrentCell { get; set; }
        /// <summary>
        /// The worksheet where the current formula is located.
        /// </summary>
        public ExcelWorksheet CurrentWorksheet 
        { 
            get
            {
                if(Package != null && CurrentCell.WorksheetIx>=0 && CurrentCell.WorksheetIx < Package.Workbook.Worksheets.Count)
                {
                    return Package.Workbook.GetWorksheetByIndexInList(CurrentCell.WorksheetIx);
                }
                return null;
            }
        }
    }
}
