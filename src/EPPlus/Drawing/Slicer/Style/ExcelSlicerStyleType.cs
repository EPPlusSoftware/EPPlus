/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Slicer.Style
{
    /// <summary>
    /// A type specifing the type of style element for a named custom slicer style.
    /// </summary>
    public enum eSlicerStyleElement
    {
        /// <summary>
        /// Styles a slicer item with data that is not selected
        /// </summary>
        UnselectedItemWithData,
        /// <summary>
        /// Styles a slicer item that is selected
        /// </summary>
        SelectedItemWithData,
        /// <summary>
        /// Styles a slicer item with no data that is not selected
        /// </summary>
        UnselectedItemWithNoData,
        /// <summary>
        /// Styles a select slicer item with no data.
        /// </summary>
        SelectedItemWithNoData,
        /// <summary>
        /// Styles a slicer item with data that is not selected and over which the mouse is paused on
        /// </summary>
        HoveredUnselectedItemWithData,
        /// <summary>
        /// Styles a selected slicer item with data and over which the mouse is paused on
        /// </summary>
        HoveredSelectedItemWithData,
        /// <summary>
        /// Styles a slicer item with no data that is not selected and over which the mouse is paused on
        /// </summary>
        HoveredUnselectedItemWithNoData,
        /// <summary>
        /// Styles a selected slicer item with no data and over which the mouse is paused on
        /// </summary>
        HoveredSelectedItemWithNoData
    }
}
