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
    public class TextFragment : TextFragmentBase
    {
        MeasurementFont _mfFont;

        public MeasurementFont Font { get { return _mfFont; } set { _mfFont = value; base.RichTextOptions.SetFont(value); } }

        public TextFragment(IRichTextInfoBase rtFormat) : base(rtFormat)
        {
        }
        public TextFragment():base()
        {

        }

        ///// <summary>
        ///// Store rich-text info.
        ///// Nothing is supposed to be done with this within OpenType
        ///// but we hold the data so users may more easily recognize what rich text this is in the output.
        ///// </summary>
        //public new IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();

        public override float Size { get => Font.Size; }
    }

    public class TextFragmentBase
    {
        public string Text { get => RichTextOptions.Text; set => RichTextOptions.Text = value; }

        public IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }

        public TextFragmentBase()
        {
        }
        public TextFragmentBase(IRichTextInfoBase richText) 
        {
            RichTextOptions = richText;
        }
        public virtual float Size { get => RichTextOptions.Size; }
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
