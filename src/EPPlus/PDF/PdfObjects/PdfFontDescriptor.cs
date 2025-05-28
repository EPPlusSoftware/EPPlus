using OfficeOpenXml.PDF.PdfPageSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    [Flags]
    public enum FontDescriptorFlags
    {
        FixedPitch = 1,
        Serif = 2,
        Symbolic = 4,
        Script = 8,
        NonSymbolic = 32,
        Italic = 64,
        AllCap = 65536,
        SmallCap = 131072,
        ForceBold = 262144,
    }

    internal class PdfFontDescriptor : PdfObject
    {
        private readonly string fontName; //Same as PdfFont.BaseFont
        private readonly int flags;
        private readonly PdfRect fontBBox;
        private readonly double italicAngle;
        private readonly int ascent;
        private readonly int descent;
        private readonly double stemV;
        private readonly int capheight;

        public PdfFontDescriptor(int objectNumber, string fontName, int flags, PdfRect fontBBox, double italicAngle, int ascent, int descent, double stemV, int capHeight, int version = 0)
            : base(objectNumber, version)
        {
            this.fontName = fontName;
            this.flags = flags;
            this.fontBBox = fontBBox;
            this.italicAngle = italicAngle;
            this.ascent = ascent;
            this.descent = descent;
            this.stemV = stemV;
            this.capheight = capHeight;
        }

        internal override string RenderDictionary()
        {
            return $"<<  /Type /FontDescriptor\n" +
                    $"   /FontName /{fontName.Replace(" ", "")}\n" +
                    $"   /Flags {flags}\n" +
                    $"   /FontBBox [{fontBBox.X + fontBBox.Width} {fontBBox.Y} {fontBBox.X+fontBBox.Width} {fontBBox.Y + fontBBox.Height}]\n" +
                    $"   /Ascent {ascent}\n" +
                    $"   /Descent {descent}\n" +
                    $"   /CapHeight {capheight}\n" +
                    $"   /ItalicAngle {(int)italicAngle}\n" +
                    $"   /StemV {(int)stemV} >>";
        }
    }
}
