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
using EPPlus.Graphics;
using System;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    /// <summary>
    /// Font descriptor flags
    /// </summary>
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
        private readonly string fontName;
        private readonly int flags;
        private readonly Rect fontBBox;
        private readonly double italicAngle;
        private readonly int ascent;
        private readonly int descent;
        private readonly double stemV;
        private readonly int capheight;



        public PdfFontDescriptor(int objectNumber, string fontName, int flags, Rect fontBBox, double italicAngle, int ascent, int descent, double stemV, int capHeight, int version = 0)
            : base(objectNumber, version)
        {
            this.fontName = fontName;
            this.flags = flags;
            this.fontBBox = fontBBox;
            this.italicAngle = italicAngle;
            this.ascent = ascent;
            this.descent = descent;
            this.stemV = stemV;
            capheight = capHeight;
        }

        internal override string RenderDictionary()
        {
            return $"<<  /Type /FontDescriptor\n" +
                    $"   /FontName /{fontName.Replace(" ", "")}\n" +
                    $"   /Flags {flags}\n" +
                    $"   /FontBBox [{fontBBox.X} {fontBBox.Y} {fontBBox.Width} {fontBBox.Height}]\n" +
                    $"   /Ascent {ascent}\n" +
                    $"   /Descent {descent}\n" +
                    $"   /CapHeight {capheight}\n" +
                    $"   /ItalicAngle {(int)italicAngle}\n" +
                    $"   /StemV {(int)stemV} >>";
        }
    }
}
