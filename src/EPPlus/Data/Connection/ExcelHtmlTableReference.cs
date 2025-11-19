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
namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// A reference to an HTML table in a web query connection.
    /// </summary>
    public class ExcelHtmlTableReference
    {
        /// <summary>
        /// Represents a missing table (Index = -1).
        /// </summary>
        public ExcelHtmlTableReference()
        {
            Index = -1;
        }
        /// <summary>
        /// Represents a table with specified index within the html document.
        /// </summary>
        /// <param name="index">The index</param>
        public ExcelHtmlTableReference(int index)
        {
            Index = index;
        }
        /// <summary>
        /// Represents a table with the specified name in the html document using the "o:WebQueryTableName" attribute.
        /// </summary>
        /// <param name="name">The name.</param>
        public ExcelHtmlTableReference(string name)
        {
            Name=name;
            Index = 0;
        }
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