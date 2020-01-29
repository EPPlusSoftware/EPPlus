/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.Font;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Interfaces
{
    /// <summary>
    /// Interface to handle styles on a chart part
    /// </summary>
    internal interface IDrawingStyleBase
    {
        /// <summary>
        /// Create the spPr element within the drawing part if does not exist.
        /// </summary>
        void CreatespPr();
        /// <summary>
        /// Border settings
        /// </summary>
        ExcelDrawingBorder Border { get; }
        /// <summary>
        /// Effect settings
        /// </summary>
        ExcelDrawingEffectStyle Effect { get; }
        /// <summary>
        /// Fill settings
        /// </summary>
        ExcelDrawingFill Fill { get; }
        /// <summary>
        /// 3D settings
        /// </summary>
        ExcelDrawing3D ThreeD { get; }
    }
    /// <summary>
    /// Interface to handle font styles on a chart part
    /// </summary>
    internal interface IDrawingStyle : IDrawingStyleBase
    {
        /// <summary>
        /// Font settings
        /// </summary>
        ExcelTextFont Font { get;  }
        /// <summary>
        /// Text body settings
        /// </summary>
        ExcelTextBody TextBody { get; }
    }
    internal interface IStyleMandatoryProperties
    {
        void SetMandatoryProperties();
    }
}
