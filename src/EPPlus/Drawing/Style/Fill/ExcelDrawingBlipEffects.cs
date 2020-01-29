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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// Effects added to a blip fill
    /// </summary>
    public class ExcelDrawingBlipEffects : XmlHelper
    {
        private const string  _duoTonePath = "a:duotone";
        private const string _clrChangePath = "a:clrChange";
        private const string _clrReplPath = "a:clrRepl";
        internal ExcelDrawingBlipEffects(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            SchemaNodeOrder = new string[]
            {
                "alphaBiLevel",
                "alphaCeiling",
                "alphaFloor",
                "alphaInv",
                "alphaMod",
                "alphaModFix",
                "alphaRepl",
                "biLevel",
                "blur",
                "clrChange",
                "clrRepl",
                "duotone",
                "fillOverlay",
                "grayscl",
                "hsl",
                "lum",
                "tint"
            };

            var node = GetNode(_duoTonePath);
            if(node!=null)
            {
                Duotone = new ExcelDrawingDuotoneEffect(NameSpaceManager, node);
            }

            node = GetNode(_clrChangePath);
            if (node != null)
            {
                ColorChange = new ExcelDrawingColorChangeEffect(NameSpaceManager, node);
            }

            node = GetNode(_clrReplPath);
            if (node != null)
            {
                ColorReplace = new ExcelDrawingColorReplaceEffect(NameSpaceManager, node);
            }
        }
        #region DuoNode
        /// <summary>
        /// Adds a duotone effect 
        /// </summary>
        public void AddDuotone()
        {
            if(Duotone!=null)
            {
                throw (new InvalidOperationException("Duotone property is already added"));
            }
            var node = CreateNode(_duoTonePath);
            node.InnerXml = "<a:schemeClr val=\"accent1\"><a:shade val=\"36000\"/><a:satMod val=\"120000\" /></a:schemeClr><a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr>";
            Duotone = new ExcelDrawingDuotoneEffect(NameSpaceManager, node);
        }
        /// <summary>
        /// Removes a duotone effect.
        /// </summary>
        public void RemoveDuotone()
        {
            DeleteNode(_duoTonePath);
            Duotone = null;
        }
        /// <summary>
        /// A duo tone color effect.
        /// </summary>
        public ExcelDrawingDuotoneEffect Duotone 
        {
            get;
            private set;
        }
        #endregion
        #region ColorChange
        /// <summary>
        /// Adds a color change effect 
        /// </summary>
        public void AddColorChange()
        {
            if (ColorChange != null)
            {
                throw (new InvalidOperationException("ColorChange property is already added"));
            }
            var node = CreateNode(_clrChangePath);
            node.InnerXml = "<a:schemeClr val=\"accent1\"><a:shade val=\"36000\"/><a:satMod val=\"120000\" /></a:schemeClr><a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr>";
            ColorChange = new ExcelDrawingColorChangeEffect(NameSpaceManager, node);
        }
        /// <summary>
        /// Removes a duotone effect.
        /// </summary>
        public void RemoveColorChange()
        {
            DeleteNode(_clrChangePath);
            ColorChange = null;
        }
        /// <summary>
        /// A duo tone color effect.
        /// </summary>
        public ExcelDrawingColorChangeEffect ColorChange
        {
            get;
            private set;
        }
        #endregion        
        #region ColorReplace
        /// <summary>
        /// Adds a color change effect 
        /// </summary>
        public void AddColorReplace()
        {
            if (ColorReplace != null)
            {
                throw (new InvalidOperationException("ColorChange property is already added"));
            }
            var node = CreateNode(_clrReplPath);
            node.InnerXml = "<a:schemeClr val=\"accent1\" />";
            ColorReplace = new ExcelDrawingColorReplaceEffect(NameSpaceManager, node);
        }
        /// <summary>
        /// Removes a duotone effect.
        /// </summary>
        public void RemoveColorReplace()
        {
            DeleteNode(_clrReplPath);
            ColorReplace = null;
        }
        /// <summary>
        /// Adds color replacement effect.
        /// </summary>
        public ExcelDrawingColorReplaceEffect ColorReplace
        {
            get;
            private set;
        }
        #endregion        
    }
}
