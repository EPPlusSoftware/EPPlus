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
    /// <summary>
    /// Represents a text fragment with specific font properties.
    /// </summary>
    public class TextFragment : TextFragmentBase, ITextFragmentBase
    {
        /// <summary>
        /// Legacy. This is to be replaced after PDF refactor is taken in
        /// </summary>
        public IFontFormatBase Font { get { return RichTextOptions; } set {RichTextOptions.SetFont(value); } }

        public TextFragment(IRichTextFormatSimple rtFormat) : base(rtFormat)
        {
            RichTextOptions = rtFormat;
        }
        public TextFragment():base()
        {
            RichTextOptions = new RichTextFormatSimple();
        }

        public new IRichTextFormatSimple RichTextOptions { get { return (IRichTextFormatSimple)base.RichTextOptions; } set { base.RichTextOptions = value; } }

        public override float Size { get => RichTextOptions.Size; }
    }

    public class TextFragmentBase : ITextFragmentBase
    {
        public string Text { get => RichTextOptions.Text; set => RichTextOptions.Text = value; }
        /// <summary>
        /// Store rich-text info.
        /// We must extract font info from this but nothing else is supposed to be done with this within opentype
        /// but we hold the data so users may more easily recognize which rich text this is in the output.
        /// </summary>
        public virtual IRichTextFormatEssential RichTextOptions { get; set; } = new RichTextFormatBase();
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }

        public string FullFontName
        {
            get
            {
                return $"{RichTextOptions.Family} {RichTextOptions.SubFamily.ToString().Replace(", ", " ")}";
            }
        }

        public TextFragmentBase()
        {
        }
        public TextFragmentBase(IRichTextFormatEssential richText) 
        {
            RichTextOptions = richText;
        }
        public virtual float Size { get => RichTextOptions.Size; }
    }
}
