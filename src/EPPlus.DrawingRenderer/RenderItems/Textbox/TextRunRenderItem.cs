using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using System.Drawing;
using System.Text.RegularExpressions;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    /// <summary>
    /// Linestyle
    /// </summary>
    public enum UnderLineType
    {
        /// <summary>
        /// Dashed
        /// </summary>
        Dash,
        /// <summary>
        /// Dashed, Thicker
        /// </summary>
        DashHeavy,
        /// <summary>
        /// Dashed Long
        /// </summary>
        DashLong,
        /// <summary>
        /// Long Dashed, Thicker
        /// </summary>
        DashLongHeavy,
        /// <summary>
        /// Double lines with normal thickness
        /// </summary>
        Double,
        /// <summary>
        /// Dot Dash
        /// </summary>
        DotDash,
        /// <summary>
        /// Dot Dash, Thicker
        /// </summary>
        DotDashHeavy,
        /// <summary>
        /// Dot Dot Dash
        /// </summary>
        DotDotDash,
        /// <summary>
        /// Dot Dot Dash, Thicker
        /// </summary>
        DotDotDashHeavy,
        /// <summary>
        /// Dotted
        /// </summary>
        Dotted,
        /// <summary>
        /// Dotted, Thicker
        /// </summary>
        DottedHeavy,
        /// <summary>
        /// Single line, Thicker
        /// </summary>
        Heavy,
        /// <summary>
        /// No underline
        /// </summary>
        None,
        /// <summary>
        /// Single line
        /// </summary>
        Single,
        /// <summary>
        /// A single wavy line
        /// </summary>
        Wavy,
        /// <summary>
        /// A double wavy line
        /// </summary>
        WavyDbl,
        /// <summary>
        /// A single wavy line, Thicker
        /// </summary>
        WavyHeavy,
        /// <summary>
        /// Underline just the words and not the spaces between them
        /// </summary>
        Words
    }
    /// <summary>
    /// BulletType of font strike
    /// </summary>
    public enum StrikeType
    {
        /// <summary>
        /// Double-lined font strike
        /// </summary>
        Double,
        /// <summary>
        /// No font strike
        /// </summary>
        No,
        /// <summary>
        /// Single-lined font strike
        /// </summary>
        Single
    }
    public abstract class TextRunRenderItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.TextRun;

        protected string _originalText;
        public string _currentText { get; protected set; }

        public IFontFormatBase _measurementFont { get; internal protected set; }
        protected bool _isFirstInParagraph;

        public double FontSizeInPixels { get;  protected set; }

        public List<string> Lines { get;  set; }

        protected internal bool _isItalic = false;
        protected internal bool _isBold = false;
        protected internal UnderLineType _underLineType;
        protected internal StrikeType _strikeType;
        protected internal Color _underlineColor;
        protected internal double _baseline;

        public double YPosition { get; set; }
        public double ClippingHeight { get; protected set; } = double.NaN;
        public TextRunRenderItem(BoundingBox parent) : base(parent)
        {
            Bounds.Name = "TextRun";
        }

        public TextRunRenderItem(BoundingBox parent, IFontFormatBase font, string displayText) : base(parent)
        {
            _originalText = displayText;

            Bounds.Name = "TextRun";
            _currentText = displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            _measurementFont = font;
            _isFirstInParagraph = true;

            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);
            Bounds.Height = _measurementFont.Size;
            if (parent.Height < _measurementFont.Size)
            {
                parent.Height = _measurementFont.Size;
            }

            //To get clipping height we need to get the textbody bounds
            if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            {
                ClippingHeight = parent.Parent.Parent.Position.Y + parent.Parent.Parent.Size.Y;
            }

            if (Lines.Count == 1)
            {
                //Bounds.Width = parent.Width;
                GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            }
            else
            {
                //Measure text.
                GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            }

            _underLineType = UnderLineType.None;
        }
        public TextRunRenderItem(BoundingBox parent, string text, IFontFormatBase font, string displayText) : base(parent)
        {
            _originalText = text;

            Bounds.Name = "TextRun";
            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            _measurementFont = font;
            _isFirstInParagraph = true;
        }
    }
}
