/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing
{
    public enum eBulletSizeType
    {
        /// <summary>
        /// The size of the bullet characters is in percentage of the surrounding text within the paragraph
        /// </summary>
        PercentOfText,
        /// <summary>
        /// The size of the bullet characters is points
        /// </summary>
        Points,
        /// <summary>
        /// The size of the bullet characters is the same points as the surrounding text.
        /// </summary>
        FollowText
    }
}