/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/22/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Controls;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides a simple way to type cast control drawing object top its top level class.
    /// </summary>
    public class ExcelControlAsType
    {
        ExcelDrawing _drawing;
        internal ExcelControlAsType(ExcelDrawing drawing)
        {
            _drawing = drawing;
        }
        /// <summary>
        /// Converts the drawing to it's top level or other nested drawing class.        
        /// </summary>
        /// <typeparam name="T">The type of drawing. T must be inherited from ExcelDrawing</typeparam>
        /// <returns>The drawing as type T</returns>
        public T Type<T>() where T : ExcelControl
        {
            return _drawing as T;
        }

        /// <summary>
        /// Returns the drawing as a button. 
        /// If this drawing is not a button, null will be returned
        /// </summary>
        /// <returns>The drawing as a button</returns>
        public ExcelControlButton Button
        {
            get { return _drawing as ExcelControlButton; }
        }
        /// <summary>
        /// Returns the drawing as a drop-down. 
        /// If this drawing is not a drop-down, null will be returned
        /// </summary>
        /// <returns>The drawing as a drop-down</returns>        
        public ExcelControlDropDown DropDown
        {
            get { return _drawing as ExcelControlDropDown; }
        }
        /// <summary>
        /// Returns the drawing as a group box. 
        /// If this drawing is not a group box, null will be returned
        /// </summary>
        /// <returns>The drawing as a group box</returns>        
        public ExcelControlGroupBox GroupBox
        {
            get { return _drawing as ExcelControlGroupBox; }
        }
        /// <summary>
        /// Returns the drawing as a label. 
        /// If this drawing is not a label, null will be returned
        /// </summary>
        /// <returns>The drawing as a label</returns>        
        public ExcelControlLabel Label
        {
            get { return _drawing as ExcelControlLabel; }
        }

        /// <summary>
        /// Returns the drawing as a list box. 
        /// If this drawing is not a list box, null will be returned
        /// </summary>
        /// <returns>The drawing as a list box</returns>        
        public ExcelControlListBox ListBox
        {
            get { return _drawing as ExcelControlListBox; }
        }

        /// <summary>
        /// Returns the drawing as a check box. 
        /// If this drawing is not a check box, null will be returned
        /// </summary>
        /// <returns>The drawing as a check box</returns>        
        public ExcelControlCheckBox CheckBox
        {
            get { return _drawing as ExcelControlCheckBox; }
        }

        /// <summary>
        /// Returns the drawing as a radio button. 
        /// If this drawing is not a radio button, null will be returned
        /// </summary>
        /// <returns>The drawing as a radio button</returns>        
        public ExcelControlRadioButton RadioButton
        {
            get { return _drawing as ExcelControlRadioButton; }
        }

        /// <summary>
        /// Returns the drawing as a scroll bar. 
        /// If this drawing is not a scroll bar, null will be returned
        /// </summary>
        /// <returns>The drawing as a scroll bar</returns>        
        public ExcelControlScrollBar ScrollBar
        {
            get { return _drawing as ExcelControlScrollBar; }
        }

        /// <summary>
        /// Returns the drawing as a spin button. 
        /// If this drawing is not a spin button, null will be returned
        /// </summary>
        /// <returns>The drawing as a spin button</returns>        
        public ExcelControlSpinButton SpinButton
        {
            get { return _drawing as ExcelControlSpinButton; }
        }
    }
}
