/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/22/2025         EPPlus Software AB       EPPlus 8.0.8
 *************************************************************************************************/
namespace OfficeOpenXml.VBA
{

    /// <summary>
    /// The interface that must be implemented by the elements stored by ExcelVBACollectionBase.
    /// </summary>
    public interface IExcelVBACollectionElement
    {
        string Name { get; }
    }
}
