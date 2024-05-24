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
    /// Field for database criteria
    /// </summary>
    internal class ExcelDatabaseCriteriaField
    {
        /// <summary>
        /// Constructor with field name
        /// </summary>
        /// <param name="fieldName"></param>
        public ExcelDatabaseCriteriaField(string fieldName)
        {
            FieldName = fieldName;
        }
        /// <summary>
        /// Constructor with field index
        /// </summary>
        /// <param name="fieldIndex"></param>
        public ExcelDatabaseCriteriaField(int fieldIndex)
        {
            FieldIndex = fieldIndex;
        }
        /// <summary>
        /// return name or object toString
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (!string.IsNullOrEmpty(FieldName))
            {
                return FieldName;
            }
            return base.ToString();
        }
        /// <summary>
        /// Name of field
        /// </summary>
        public string FieldName { get; private set; }
        /// <summary>
        /// Index of field
        /// </summary>
        public int? FieldIndex { get; private set; }
    }
}
