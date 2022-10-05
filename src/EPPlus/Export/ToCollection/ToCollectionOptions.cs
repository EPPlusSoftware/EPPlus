/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/04/2022         EPPlus Software AB       Initial release EPPlus 6.1
 *************************************************************************************************/
#if !NET35 && !NET40
#endif
namespace OfficeOpenXml
{
    /// <summary>
    /// Parameters for the ToCollection Method
    /// </summary>
    public class ToCollectionOptions
    {
        ///// <summary>
        ///// The type of value returned for the cells.
        ///// </summary>
        //public ToCollectionValueType ValueType { get; set; } = ToCollectionValueType.Value;
        /// <summary>
        /// Header row in the range, if applicable. 
        /// A null value means there is no header row.
        /// See also: <seealso cref="Headers"/>
        /// <seealso cref="DataStartRow"/>
        /// </summary>
        public int? HeaderRow { get; set; } = null;
        /// <summary>
        /// Data start row in the range.
        /// A null value means, the data rows starts direcly after the header row.
        /// </summary>
        public int? DataStartRow { get; set; } = null;

        /// <summary>
        /// An array of column headers. If set, used instead of the header row. 
        /// <see cref="SetCustomHeaders(string[])"/>
        /// </summary>
        internal string[] Headers { get; private set; } = null;
        /// <summary>
        /// Sets custom headers
        /// </summary>
        /// <param name="header"></param>
        public void SetCustomHeaders(params string[] header)
        {
            Headers = header;
        }
    }
}