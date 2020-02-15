using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.DataValidation.Events
{
    internal class ValidationFormulaChangedArgs : EventArgs
    {
        public string ValidationUid { get; set; }

        public string OldValue { get; set; }

        public string NewValue { get; set; }
    }
}
