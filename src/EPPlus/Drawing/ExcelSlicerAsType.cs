/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  06/05/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Slicer;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides easy type cast for slicer drawings.
    /// </summary>
    public class ExcelSlicerAsType
    {
        ExcelDrawing _drawing;
        internal ExcelSlicerAsType(ExcelDrawing drawing)
        {
            _drawing = drawing;
        }
        /// <summary>
        /// Returns the drawing as table slicer . 
        /// If this drawing is not a table slicer, null will be returned
        /// </summary>
        /// <returns>The drawing as a table slicer</returns>
        public ExcelTableSlicer TableSlicer
        {
            get
            {
                return _drawing as ExcelTableSlicer;
            }
        }
        /// <summary>
        /// Returns the drawing as pivot table slicer . 
        /// If this drawing is not a pivot table slicer, null will be returned
        /// </summary>
        /// <returns>The drawing as a pivot table slicer</returns>
        public ExcelPivotTableSlicer PivotTableSlicer
        {
            get
            {
                return _drawing as ExcelPivotTableSlicer;
            }
        }
    }
}