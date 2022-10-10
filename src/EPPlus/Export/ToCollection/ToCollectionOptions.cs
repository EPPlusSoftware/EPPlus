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
using OfficeOpenXml.Export.ToCollection;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// Parameters for the ToCollection Method
    /// </summary>
    public class ToCollectionRangeOptions : ToCollectionOptions
    {
        public ToCollectionRangeOptions()
        {

        }
        internal ToCollectionRangeOptions(ToCollectionOptions options)
        {
            SetCustomHeaders(options.Headers);
        }
        /// <summary>
        /// Header row in the range, if applicable. 
        /// A null value means there is no header row.
        /// See also: <seealso cref="ToCollectionOptions.SetCustomHeaders(string[])"/>
        /// <seealso cref="DataStartRow"/>
        /// </summary>
        public int? HeaderRow { get; set; } = null;
        /// <summary>
        /// The data start row in the range.
        /// A null value means the data rows starts direcly after the header row.
        /// </summary>
        public int? DataStartRow { get; set; } = null;
        /// <summary>
        /// A <see cref="ToCollectionRangeOptions"/> with default values.
        /// </summary>
        public static ToCollectionRangeOptions Default
        {
            get
            {
                return new ToCollectionRangeOptions();
            }
        }
    }
    public abstract class ToCollectionOptions
    {
        /// <summary>
        /// An array of column headers. If set, used instead of the header row. 
        /// <see cref="SetCustomHeaders(string[])"/>
        /// </summary>
        internal string[] Headers { get; private set; } = null;
        /// <summary>
        /// Sets custom headers.  If set, used instead of the header row. 
        /// </summary>
        /// <param name="header"></param>
        public void SetCustomHeaders(params string[] header)
        {
            Headers = header;
        }
        /// <summary>
        /// How conversion failures should be handled when mapping properties.
        /// </summary>
        public ToCollectionConversionFailureStrategy ConversionFailureStrategy { get; set; }
    }
    /// <summary>
    /// Table ToCollection options.
    /// </summary>
    public class ToCollectionTableOptions : ToCollectionOptions
    {
        /// <summary>
        /// A <see cref="ToCollectionTableOptions"/> with default values.
        /// </summary>
        public static ToCollectionTableOptions Default
        {
            get
            {
                return new ToCollectionTableOptions();
            }
        }
    }
}