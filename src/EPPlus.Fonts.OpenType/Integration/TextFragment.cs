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
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System.Drawing;

namespace EPPlus.Fonts.OpenType.Integration
{

    public interface IFragInfo
    {
        public double AscentPoints { get; }
        public double DescentPoints { get; }

        public float Size { get; }
    }

    /// <summary>
    /// Represents a text fragment with specific font properties.
    /// </summary>
    public class TextFragment : TextFragmentBase, ITextFragmentBase
    {
        /// <summary>
        /// Legacy. This is to be replaced after PDF refactor is taken in
        /// </summary>
        MeasurementFont _mfFont;

        /// <summary>
        /// Legacy. This is to be replaced after PDF refactor is taken in
        /// </summary>
        public MeasurementFont Font { get { return _mfFont; } set { _mfFont = value; RichTextOptions.SetFont(value); } }

        public TextFragment(IRichTextInfoBase rtFormat) : base(rtFormat)
        {
            RichTextOptions = rtFormat;
        }
        public TextFragment():base()
        {

        }
        /// <summary>
        /// Store rich-text info.
        /// We must extract font info from this but nothing else is supposed to be done with this within opentype
        /// </summary>
        public IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();

        public override IRichTextFormatBase RichTextFormat { get => RichTextOptions; set => RichTextOptions = (IRichTextInfoBase)value; }

        public override float Size { get => RichTextOptions.Size; }
    }

    public class TextFragmentBase : ITextFragmentBase
    {
        public string Text { get => RichTextFormat.Text; set => RichTextFormat.Text = value; }
        /// <summary>
        /// Store rich-text info.
        /// We must extract font info from this but nothing else is supposed to be done with this within opentype
        /// but we hold the data so users may more easily recognize which rich text this is in the output.
        /// </summary>
        public virtual IRichTextFormatBase RichTextFormat { get; set; } = new OpenTypeRichTextBase();
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }

        public TextFragmentBase()
        {
        }
        public TextFragmentBase(IRichTextFormatBase richText) 
        {
            RichTextFormat = richText;
        }
        public virtual float Size { get => RichTextFormat.Size; }
    }

    ///// <summary>
    ///// Simple class to provide some kind of fallback/defaults
    ///// </summary>
    //public class RichTextDefaults : IRichTextInfoBase
    //{
    //    internal RichTextDefaults()
    //    {
    //    }
    //    public bool Italic { get; set; } = false;

    //    public bool Bold { get; set; } = false;

    //    public bool SubScript { get; set; } = false;

    //    public bool SuperScript { get; set; } = false;

    //    public int UnderlineType { get; set; } = -1;

    //    public int StrikeType { get; set; } = -1;

    //    public int Capitalization { get; set; } = -1;

    //    public Color UnderlineColor { get; set; }

    //    public Color FontColor { get; set; }
    //}
}
