/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents a defined name in an external workbook
    /// </summary>
    public class ExcelExternalDefinedName : IExcelExternalNamedItem
    {
        /// <summary>
        /// The name
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// The address that the defined name referes to
        /// </summary>
        public string RefersTo { get; internal set; }
        /// <summary>
        /// The sheet id
        /// </summary>
        public int SheetId { get; internal set; }
        /// <summary>
        /// The string representation of the name
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Name;
        }
    }
}
