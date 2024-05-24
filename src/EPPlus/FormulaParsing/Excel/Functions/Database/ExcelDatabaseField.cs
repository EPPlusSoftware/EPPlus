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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    /// <summary>
    /// Database field
    /// </summary>
    internal class ExcelDatabaseField
    {
        /// <summary>
        /// Name of field
        /// </summary>
        public string FieldName { get; private set; }
        /// <summary>
        /// Column index
        /// </summary>
        public int ColIndex { get; private set; }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="colIndex"></param>
        public ExcelDatabaseField(string fieldName, int colIndex)
        {
            FieldName = fieldName;
            ColIndex = colIndex;
        }
    }
}
