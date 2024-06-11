/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/ 
  A commercial license to use this software can be purchased at https://epplussoftware.com 
  Date               Author                       Change
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/

namespace OfficeOpenXml.Core
{
    public class AutofitParams
    {
        /// <summary>
        /// The ammount of rows to check for autofitting, starts from top.
        /// </summary>
        public int Rows = 0;

        /// <summary>
        /// A percentage of the widest text. Since charaters in different fonts have different widths we use this threshold remove characters from the longer string for comparing to the current text.
        /// This is so we can skip obvious shorter strings and save time on calculating it's actual width.
        /// </summary>
        public double textLengthThreshold = 0.75d;

    }
}
