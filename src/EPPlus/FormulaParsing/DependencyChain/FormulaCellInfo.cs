using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
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

        public string Worksheet => _worksheet;

        public string Address => _address;

        public string Formula => _formula;
    }
}
