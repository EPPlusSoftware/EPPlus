using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata
{
    /// <summary>
    /// Attribute used for Excel formula functions metadata.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class FunctionMetadataAttribute : Attribute
    {

        /// <summary>
        /// Function category
        /// </summary>
        public ExcelFunctionCategory Category { get; set; }

        /// <summary>
        /// EPPlus version where the function was introduced
        /// </summary>
        public string EPPlusVersion { get; set; }

        /// <summary>
        /// Short description of the function.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// A string describing in which Excel version the function was introduced.
        /// </summary>
        public string IntroducedInExcelVersion { get; set; }

        /// <summary>
        /// Returns true if the function can return an array if called with a multicell range as the argument.
        /// </summary>
        public bool SupportsArrays { get; set; }

    }
}
