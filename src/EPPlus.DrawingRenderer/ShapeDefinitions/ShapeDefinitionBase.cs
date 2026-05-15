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
using EPPlus.DrawingRenderer.ShapeDefinitions;
using EPPlus.DrawingRenderer.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    [DebuggerDisplay("{Style}")]
    public abstract class ShapeDefinitionBase
    {
        public Dictionary<string, double> _calculatedValues = new Dictionary<string, double>();

        protected ShapeDefinitionBase()
        {
                
        }
        /// <summary>
        /// Clone constructor
        /// </summary>
        /// <param name="original">The original to clone from</param>
        protected ShapeDefinitionBase(ShapeDefinitionBase original)
        {
            Style=original.Style;
            if (original.ShapeAdjustValues != null)
            {
                ShapeAdjustValues = new List<ShapeGuide>();
                foreach (var av in original.ShapeAdjustValues)
                {
                    ShapeAdjustValues.Add(av.Clone());
                }
            }
            if (original.ShapeGuides != null)
            {
                ShapeGuides = new List<ShapeGuide>();
                foreach (var g in original.ShapeGuides)
                {
                    ShapeGuides.Add(g.Clone());
                }
            }

            if (original.ShapeAdjustHandles!=null)
            {
                ShapeAdjustHandles = new List<ShapeAdjustHandleBase>();
                foreach (var ah in original.ShapeAdjustHandles)
                {
                    ShapeAdjustHandles.Add(ah.Clone());
                }
            }

            TextBoxRect = original.TextBoxRect?.Clone();

            ShapePaths = new List<DrawingPath>();
            foreach (var p in original.ShapePaths)
            {
                ShapePaths.Add(p.Clone());
            }
        }


        public ShapeStyle Style { get; set; }
        /// <summary>
        /// avLst
        /// </summary>
        public List<ShapeGuide> ShapeAdjustValues { get; set; }
        /// <summary>
        /// gdLst  
        /// </summary>
        public List<ShapeGuide> ShapeGuides { get; set; }
        /// <summary>
        /// ahLst 
        /// </summary>
        public List<ShapeAdjustHandleBase> ShapeAdjustHandles { get; set; }
        //cxnLst 
        public List<ShapeConnectionSite> ShapeConnectionSite { get; set; }
        //rect 
        /// <summary>
        /// The rectangle for the text inside the shape.
        /// </summary>
        public TextBoxRect TextBoxRect { get; set; }
        //pathLst
        /// <summary>
        /// Paths to draw the shape
        /// </summary>
        public List<DrawingPath> ShapePaths { get; set; }


        //private object GetValueOfNameOrCalculateValue(object value)
        //{
        //    if (value is string s && _calculatedValues.ContainsKey(s))
        //    {
        //        return _calculatedValues[s];
        //    }
        //    return value;
        //}

        public abstract ShapeDefinitionBase Clone();
    }
}