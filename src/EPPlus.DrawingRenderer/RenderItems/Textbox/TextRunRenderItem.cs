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

        public int OriginalRtIdx { get; private set; } = -1;

        protected string _originalText;
        public string _currentText { get; protected set; }

        public IFontFormatBase _measurementFont { get; internal protected set; }
        protected bool _isFirstInParagraph;

        public double FontSizeInPixels { get;  protected set; }

        public List<string> Lines { get;  set; }

        protected internal bool _isItalic = false;
        protected internal bool _isBold = false;
        protected internal UnderLineType _underLineType = UnderLineType.None;
        protected internal StrikeType _strikeType;
        protected internal Color _underlineColor;
        protected internal double _baseline;

        public double YPosition { get; set; }
        public double ClippingHeight { get; protected set; } = double.NaN;
        public TextRunRenderItem(BoundingBox parent) : base(parent)
        {
            Bounds.Name = "TextRun";
        }

        public TextRunRenderItem(BoundingBox parent, string text, int origRtIdx) : base(parent)
        {
            Bounds.Name = "TextRun";
            _currentText = text;
            OriginalRtIdx = origRtIdx;
        }

        internal protected void InitializeBase(IFontFormatBase font)
        {
            //Should be ascent-only?
            Bounds.Height = font.Size;
            FontSizeInPixels = ((double)font.Size).PointToPixel(true);
            _measurementFont = font;
        }

        /// <summary>
        /// Initialization for the two lower constructors
        /// Not a Initialize() method since compiler warns of un-initalized variables if you do.
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="origText"></param>
        /// <param name="currentText"></param>
        /// <param name="font"></param>
        private TextRunRenderItem(BoundingBox parent, string origText, string currentText, IFontFormatBase font) : base(parent)
        {
            Bounds.Name = "TextRun";

            //possibly no longer neccesary
            _originalText = origText;

            _currentText = currentText;

            //Should be ascent-only?
            Bounds.Height = font.Size;

            //Possibly no longer neccesary
            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            FontSizeInPixels = ((double)font.Size).PointToPixel(true);

            _measurementFont = font;
            _isFirstInParagraph = true;
        }

        public TextRunRenderItem(BoundingBox parent, IFontFormatBase font, string displayText) 
            : this(parent, displayText, displayText, font)
        {
            ////Dash is default but we know there is no underline in our input here
            //_underLineType = UnderLineType.None;
        }
        public TextRunRenderItem(BoundingBox parent, string text, IFontFormatBase font, string displayText) 
            : this(parent, text, string.IsNullOrEmpty(displayText) ? text : displayText, font)
        {
        }

        internal protected void CalculateClippingHeightFromTextBodyParent()
        {
            //To get clipping height we need to get the textbody bounds
            if (Bounds.Parent != null && Bounds.Parent.Parent != null && Bounds.Parent.Parent.Parent != null)
            {
                ClippingHeight = Bounds.Parent.Parent.Parent.Position.Y + Bounds.Parent.Parent.Parent.Size.Y;
            }
        }
    }
}
