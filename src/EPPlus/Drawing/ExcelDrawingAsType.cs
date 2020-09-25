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
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Slicer;
using System;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides a simple way to type cast drawing object top its top level class.
    /// </summary>
    public class ExcelDrawingAsType
    {
        ExcelDrawing _drawing;
        internal ExcelDrawingAsType(ExcelDrawing drawing)
        {
            _drawing = drawing;
        }
        /// <summary>
        /// Converts the drawing to it's top level or other nested drawing class.        
        /// </summary>
        /// <typeparam name="T">The type of drawing. T must be inherited from ExcelDrawing</typeparam>
        /// <returns>The drawing as type T</returns>
        public T Type<T>() where T : ExcelDrawing
        {
            return _drawing as T;
        }
        /// <summary>
        /// Returns the drawing as a shape. 
        /// If this drawing is not a shape, null will be returned
        /// </summary>
        /// <returns>The drawing as a shape</returns>
        public ExcelShape Shape
        {
            get
            {
                return _drawing as ExcelShape;
            }
        }
        /// <summary>
        /// Returns return the drawing as a picture/image. 
        /// If this drawing is not a picture, null will be returned
        /// </summary>
        /// <returns>The drawing as a picture</returns>
        public ExcelPicture Picture
        {
            get
            {
                return _drawing as ExcelPicture;
            }
        }
        ExcelChartAsType _chartAsType;
        public ExcelChartAsType Chart
        {
            get
            {
                if (_chartAsType == null)
                {
                    _chartAsType = new ExcelChartAsType(_drawing);
                }
                return _chartAsType;
            }
        }

        ExcelSlicerAsType _slicerAsType;
        public ExcelSlicerAsType Slicer 
        { 
            get
            {
                if (_slicerAsType == null)
                {
                    _slicerAsType = new ExcelSlicerAsType(_drawing);
                }
                return _slicerAsType;
            }
        }
    }
}
