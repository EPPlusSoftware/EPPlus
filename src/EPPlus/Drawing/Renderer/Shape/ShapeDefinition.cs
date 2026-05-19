/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using EPPlusImageRenderer.ShapeDefinitions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using EPPlus.DrawingRenderer.ShapeDefinitions;
using EPPlus.DrawingRenderer;

namespace OfficeOpenXml.Drawing.Renderer.Shape
{
    [DebuggerDisplay("{Style}")]
    internal class ShapeDefinition : ShapeDefinitionBase
    {
        public ShapeDefinition() : base()
        {
                
        }
        /// <summary>
        /// Clone constructor
        /// </summary>
        /// <param name="original">The original to clone from</param>
        public ShapeDefinition(ShapeDefinition original) : base(original)
        {
        }
    }
}