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

        public string IntroducedInExcelVersion { get; set; }

    }
}
