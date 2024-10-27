/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData
{
    /// <summary>
    /// Represents a rich value in the cell store
    /// </summary>
    public abstract class RichDataReference
    {
        internal RichDataReference(uint vmId, RichDataReferenceTypes refType, bool isValueError)
        {
            VmId = vmId;
            IsValueError = isValueError;
            ReferenceType = refType;
        }

        /// <summary>
        /// Rich data Id (internal for EPPlus)
        /// </summary>
        internal uint VmId { get; private set; }

        internal bool IsValueError { get; private set; }

        /// <summary>
        /// Identifies what type of rich data being referenced
        /// </summary>
        public RichDataReferenceTypes ReferenceType { get; private set; }

        /// <summary>
        /// Returns a string that reprsents the current object
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if(IsValueError)
            {
                return ExcelErrorValue.Values.Value;
            }
            return base.ToString();
        }


    }
}
