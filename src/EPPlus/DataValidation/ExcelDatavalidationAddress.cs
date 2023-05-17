using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Handling for ExcelAdress updates of DataValidations
    /// </summary>
    public class ExcelDatavalidationAddress : ExcelAddress
    {
        ExcelDataValidation _val;

        internal ExcelDatavalidationAddress(string address, ExcelDataValidation val) : base(address) 
        {
            _val = val;
        }

        /// <summary>
        /// Called before the address changes
        /// </summary>
        internal protected override void BeforeChangeAddress()
        {
            _val._ws.DataValidations.ClearRangeDictionary(_val.Address);
        }

        /// <summary>
        /// Called when the address changes
        /// </summary>
        internal protected override void ChangeAddress()
        {
            _val._ws.DataValidations.AddToRangeDictionary(_val);
        }
    }
}
