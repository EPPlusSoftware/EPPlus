/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
namespace OfficeOpenXml.Connection
{
    /// <summary>
    /// A reference to an HTML table in a web query connection.
    /// </summary>
    public class ExcelHtmlTableReference
    {
        /// <summary>
        /// The index in order of the tables in the HTML document
        /// </summary>
        public int Index 
        { 
            get; 
            set; 
        }
        /// <summary>
        /// The name of the table, if specified in the HTML source using the "o:WebQueryTableName" attribute.   
        /// </summary>
        public string Name 
        { 
            get; 
            set;
        }
    }
}