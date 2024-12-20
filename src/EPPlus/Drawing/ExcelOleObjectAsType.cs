/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  20/12/2024         EPPlus Software AB       EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.OleObject;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides a simple way to type cast control drawing object top its top level class.
    /// </summary>
    public class ExcelOleObjectAsType
    {
        ExcelDrawing _drawing;
        internal ExcelOleObjectAsType(ExcelDrawing drawing)
        {
            _drawing = drawing;
        }
        /// <summary>
        /// Converts the drawing to it's top level or other nested drawing class.        
        /// </summary>
        /// <typeparam name="T">The type of drawing. T must be inherited from ExcelDrawing</typeparam>
        /// <returns>The drawing as type T</returns>
        public T Type<T>() where T : ExcelOleObject
        {
            return _drawing as T;
        }
    }
}
