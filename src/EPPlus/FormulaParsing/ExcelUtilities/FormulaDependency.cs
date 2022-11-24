/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class FormulaDependency
    {
        public FormulaDependency(ParsingScope scope)
	    {   
            ScopeId = scope.ScopeId;
            Address = scope.Address;
	    }
        public Guid ScopeId { get; private set; }

        public FormulaRangeAddress Address { get; private set; }

        private List<FormulaRangeAddress> _referencedBy = new List<FormulaRangeAddress>();

        private List<FormulaRangeAddress> _references = new List<FormulaRangeAddress>();

        public virtual void AddReferenceFrom(FormulaRangeAddress rangeAddress)
        {
            if (Address.CollidesWith(rangeAddress)!=ExcelAddressBase.eAddressCollition.No || _references.Exists(x => x.CollidesWith(rangeAddress)!=ExcelAddressBase.eAddressCollition.No))
            {
                throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
            }
            _referencedBy.Add(rangeAddress);
        }

        //public virtual void AddReferenceTo(RangeAddress rangeAddress)
        //{
        //    if (Address.CollidesWith(rangeAddress) != ExcelAddressBase.eAddressCollition.No || _referencedBy.Exists(x => x.CollidesWith(rangeAddress) != ExcelAddressBase.eAddressCollition.No))
        //    {
        //        throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
        //    }
        //    _references.Add(rangeAddress);
        //}
    }
}
