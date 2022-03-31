/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/29/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// How the items that are used in slicer cross filtering are displayed
    /// </summary>
    public enum eCrossFilter
    {
        /// <summary>
        /// The slicer style for slicer items with no data is not applied to slicer items with no data, and slicer items with no data are not sorted separately in the list of slicer items in the slicer view.
        /// </summary>
        None,
        /// <summary>
        /// The slicer style for slicer items with no data is applied to slicer items with no data, and slicer items with no data are sorted at the bottom in the list of slicer items in the slicer view.
        /// </summary>
        ShowItemsWithDataAtTop,
        /// <summary>
        /// The slicer style for slicer items with no data is applied to slicer items with no data, and slicer items with no data are not sorted separately in the list of slicer items in the slicer view.
        /// </summary>
        ShowItemsWithNoData
    }
}