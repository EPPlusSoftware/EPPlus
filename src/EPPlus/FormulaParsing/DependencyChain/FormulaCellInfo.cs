using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Used in the formula calculation dependency chain
    /// </summary>
    public class FormulaCellInfo : IFormulaCellInfo
    {
        internal FormulaCellInfo(string worksheet, string address, string formula)
        {
            _worksheet = worksheet;
            _address = address;
            _formula = formula;
        }

        private readonly string _worksheet;
        private readonly string _address;
        private readonly string _formula;

        /// <summary>
        /// The name of the worksheet.
        /// </summary>
        public string Worksheet => _worksheet;

        /// <summary>
        /// The address of the formula
        /// </summary>
        public string Address => _address;
        /// <summary>
        /// The formula
        /// </summary>
        public string Formula => _formula;
    }
}
