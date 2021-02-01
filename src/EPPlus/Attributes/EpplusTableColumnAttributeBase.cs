﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    public abstract class EpplusTableColumnAttributeBase : Attribute
    {

        /// <summary>
        /// Order of the columns value, default value is 0
        /// </summary>
        public int Order
        {
            get;
            set;
        }

        /// <summary>
        /// Name shown in the header row, overriding the property name
        /// </summary>
        public string Header
        {
            get;
            set;
        }

        /// <summary>
        /// Excel format string for the column
        /// </summary>
        public string NumberFormat
        {
            get;
            set;
        }

        /// <summary>
        /// If not <see cref="RowFunctions.None"/> the last cell in the column (the totals row) will contain a formula of the specified type.
        /// </summary>
        public RowFunctions TotalsRowFunction
        {
            get;
            set;
        } = RowFunctions.None;

        /// <summary>
        /// Formula for the total row of this column.
        /// </summary>
        public string TotalsRowFormula
        {
            get;
            set;
        }

        /// <summary>
        /// Number format for this columns cell in the totals row.
        /// </summary>
        public string TotalsRowNumberFormat
        {
            get;
            set;
        }

        /// <summary>
        /// Text in this columns cell in the totals row
        /// </summary>
        public string TotalsRowLabel
        {
            get;
            set;
        }
    }
}
