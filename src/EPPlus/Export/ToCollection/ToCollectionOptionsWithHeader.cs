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
#if !NET35 && !NET40
#endif
namespace OfficeOpenXml
{
    public class ToCollectionOptionsWithHeader : ToCollectionOptions
    { 
        /// <summary>
        /// Header row in the range, if applicable.
        /// </summary>
        public int HeaderRow { get; set; } = 0;
        /// <summary>
        /// Data start row in the range.
        /// </summary>
        public int DataStartRow { get; set; } = 1;
    }
}