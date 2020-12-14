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

        public ExcelControlButton Button
        {
            get { return _drawing as ExcelControlButton; }
        }

        public ExcelControlDropDown DropDown
        {
            get { return _drawing as ExcelControlDropDown; }
        }
        public ExcelControlGroupBox GroupBox
        {
            get { return _drawing as ExcelControlGroupBox; }
        }
        public ExcelControlLabel Label
        {
            get { return _drawing as ExcelControlLabel; }
        }

        public ExcelControlListBox ListBox
        {
            get { return _drawing as ExcelControlListBox; }
        }

        public ExcelControlCheckBox CheckBox
        {
            get { return _drawing as ExcelControlCheckBox; }
        }

        public ExcelControlRadioButton RadioButton
        {
            get { return _drawing as ExcelControlRadioButton; }
        }

        public ExcelControlScrollBar ScrollBar
        {
            get { return _drawing as ExcelControlScrollBar; }
        }

        public ExcelControlSpinButton Spin
        {
            get { return _drawing as ExcelControlSpinButton; }
        }
    }
}
