using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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
        ExcelAddress addressBeforeChange;

        internal ExcelDatavalidationAddress(string address, ExcelDataValidation val) : base(address) 
        {
            _val = val;
            var splitAddresses = address.Split(new char[] { ' ' });
            if(splitAddresses.Count() > 1)
            {
                if(_addresses == null)
                {
                    _addresses = new List<ExcelAddressBase>();
                }

                foreach(var sAddress in splitAddresses)
                {
                    Addresses.Add(new ExcelAddress(sAddress));
                }
            }
        }

        /// <summary>
        /// Called before the address changes
        /// </summary>
        internal protected override void BeforeChangeAddress(string value)
        {
            addressBeforeChange = _val.Address;
            _val._ws.DataValidations.ClearRangeDictionary(_val.Address);
        }

        /// <summary>
        /// Called when the address changes
        /// </summary>
        internal protected override void ChangeAddress(string value)
        {
            _val._ws.DataValidations.dvQuadTree.UpdateAddress(addressBeforeChange, _val.Address, _val);
            _val._ws.DataValidations.AddToRangeDictionary(_val);
        }
    }
}
