using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.DataValidation
{
    internal class ExcelDatavalidationAddress : ExcelAddress
    {
        ExcelDataValidation _val;

        public ExcelDatavalidationAddress(string address, ExcelDataValidation val) : base(address) 
        {
            _val = val;
        }

        internal protected override void BeforeChangeAddress()
        {
            _val._ws.DataValidations.DeleteRangeDictionary(_val.Address, false);
        }

        /// <summary>
        /// Called when the address changes
        /// </summary>
        internal protected override void ChangeAddress()
        {

            _val._ws.DataValidations.UpdateRangeDictionary(_val);
        }
    }
}
