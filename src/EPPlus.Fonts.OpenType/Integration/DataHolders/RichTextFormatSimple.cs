using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    /// <summary>
    /// Defaults class mainly intended for Cells
    /// </summary>
    public class RichTextFormatSimple : RichTextFormatBase, IRichTextFormatSimple
    {
        public RichTextFormatSimple() : base()
        {

        }

        public RichTextFormatSimple(string text, string fontFamily, float size, bool bold = false, bool italic = false) : base(text, fontFamily, size, bold, italic)
        {

        }

        public bool SubScript { get; set; } = false;

        public bool SuperScript { get; set; } = false;

        public int UnderlineType { get; set; } = -1;

        public int StrikeType { get; set; } = -1;

        public int Capitalization { get; set; } = -1;

        public Color UnderlineColor { get; set; }

        public Color FontColor { get; set; }
    }
}
