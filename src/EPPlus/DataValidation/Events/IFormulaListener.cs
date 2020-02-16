using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.DataValidation.Events
{
    internal interface IFormulaListener
    {
        void Notify(ValidationFormulaChangedArgs e);
    }
}
