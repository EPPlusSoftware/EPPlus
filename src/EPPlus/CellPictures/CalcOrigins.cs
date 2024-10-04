/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.CellPictures
{
    internal enum CalcOrigins
    {
        None = 0,
        /// <summary>
        /// RichValue created directly by formula (ex, =IMAGE)
        /// </summary>
        Formula = 1,
        ComplexFormula = 2,
        DotNotation = 3,
        Reference = 4,
        /// <summary>
        /// Standalone RichValue directly stored in a cell without formula dependency (copy/paste as value or LocalImageValue)
        /// </summary>
        StandAlone = 5,
        /// <summary>
        /// Standalone RichValue created from the alt text pane after selecting "decorative"
        /// </summary>
        StandaloneDecorative = 6,
        Nested = 7,
        JSApi = 8,
        PythonResult = 9,
        Max = PythonResult
    }
}
