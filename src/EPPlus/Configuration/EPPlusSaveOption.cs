using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Configuration
{
    /// <summary>
    /// Options used when saving an <see cref="ExcelPackage"/>.
    /// Supplied via an action to the Save and SaveAs methods.
    /// </summary>
    public class EPPlusSaveOption
    {
        /// <summary>
        /// The password to encrypt the workbook with. 
        /// This overrides the Encryption.Password on the package. 
        /// If null, the password set on the package is used.
        /// </summary>
        public string Password { get; set; } = null;

        /// <summary>
        /// If true, the package is saved as an Excel template (.xltx, or .xltm if the workbook 
        /// contains a VBA project) instead of a standard workbook. Default is false.
        /// </summary>
        public bool SaveAsTemplate { get; set; } = false;
    }
}
