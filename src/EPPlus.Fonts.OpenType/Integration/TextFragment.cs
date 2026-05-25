/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextLayoutEngine implementation
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Interfaces.RichText.Interfaces;
using System.Drawing;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Represents a text fragment with specific font properties.
    /// </summary>
    public class TextFragment : ITextFragment
    {
        public string Text { get; set; }

        /// <summary>
        /// Store rich-text info.
        /// Nothing is supposed to be done with this within OpenType
        /// but we hold the data so users may more easily recognize what rich text this is in the output.
        /// </summary>
        public IRichTextInfoEssential RichText { get; } = new RtDataBasic("", "Archivo Narrow", 11f);
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }

        /// <summary>
        /// Below is to be refactored to only use RichText variable
        /// </summary>
        MeasurementFont _font { get; set; }
        public MeasurementFont Font { get { return _font; } set { _font = value; RichText.SetFont(new FontDataBasic(value)); } }
    }

    /// <summary>
    /// Minimum required richText Data + the Text it applies to.
    /// </summary>
    public class BasicTextFragment : ITextFragment
    {
        //Rich text info
        public string Text { get; set; }
        public IRichTextInfoEssential RichText { get; set; } = new RtDataBasic("", "Archivo Narrow", 11f);

        #region Measuring Info
        //Input options
        public ShapingOptions Options { get; set; }

        //Output data
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }
        #endregion
    }

    public class AdvancedTextFragment
    {
        public string Text { get => RtDataBasic.Text; set => RtDataBasic.Text = value; }
        ///// <summary>
        ///// Store rich-text info.
        ///// We must extract font info from this but nothing else is supposed to be done with this within opentype
        ///// </summary>
        //public IRichText RichText { get; }
        
        RtDataBasic RtDataBasic { get; set; }
        
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get => RtDataBasic.AscentPoints; set => RtDataBasic.DescentPoints = value; }
        public double DescentPoints { get => RtDataBasic.DescentPoints; set=> RtDataBasic.DescentPoints = value; }
    }
}
