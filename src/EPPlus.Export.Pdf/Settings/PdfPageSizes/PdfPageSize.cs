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
using EPPlus.Graphics.Units;

namespace EPPlus.Export.Pdf.Settings.PdfPageSizes
{
    public class PdfPageSize
    {
        public double Width { get; }
        public double Height { get; }
        public double WidthPu { get; }
        public double HeightPu { get; }

        public PdfPageSize(double width, double height)
        {
            Width = width;
            Height = height;
            WidthPu = System.Math.Round( UnitConversion.MmToPoints(width));
            HeightPu = System.Math.Round( UnitConversion.MmToPoints(height));
        }

        public static PdfPageSize A5 => new PdfPageSize(148d, 210d);
        public static PdfPageSize A4 => new PdfPageSize(210d, 297d); //(595, 842);
        public static PdfPageSize A3 => new PdfPageSize(297d, 420d); //(842, 1191);
        public static PdfPageSize B5 => new PdfPageSize(182d, 257d);
        public static PdfPageSize B4 => new PdfPageSize(257d, 364d);
        public static PdfPageSize Letter => new PdfPageSize(215.9d, 279.4d); //(612, 792);
        public static PdfPageSize LetterSmall => new PdfPageSize(215.9d, 279.4d);
        public static PdfPageSize Legal => new PdfPageSize(215.9d, 355.6d); //(612, 1008);
        public static PdfPageSize Statement => new PdfPageSize(139.7d, 215.9d);
        public static PdfPageSize Executive => new PdfPageSize(184.2d, 266.7d);
        public static PdfPageSize Tabloid => new PdfPageSize(279.4d, 431.8d);
        public static PdfPageSize Ledger => new PdfPageSize(431.84d, 279.4d);
        public static PdfPageSize A4Small => new PdfPageSize(210d, 297d);
        public static PdfPageSize Folio => new PdfPageSize(215.9d, 330.2d);       // 8.5 x 13 in
        public static PdfPageSize Quarto => new PdfPageSize(215d, 275d);
        public static PdfPageSize Standard10_14 => new PdfPageSize(254d, 355.6d);         // 10 x 14 in
        public static PdfPageSize Standard11_17 => new PdfPageSize(279.4d, 431.8d);       // 11 x 17 in
        public static PdfPageSize Note => new PdfPageSize(215.9d, 279.4d);       // 8.5 x 11 in
        public static PdfPageSize Envelope9 => new PdfPageSize(98.425d, 225.425d);    // 3.875 x 8.875 in
        public static PdfPageSize Envelope10 => new PdfPageSize(104.775d, 241.3d);     // 4.125 x 9.5 in
        public static PdfPageSize Envelope11 => new PdfPageSize(114.3d, 263.525d);     // 4.5 x 10.375 in
        public static PdfPageSize Envelope12 => new PdfPageSize(120.65d, 279.4d);      // 4.75 x 11 in
        public static PdfPageSize Envelope14 => new PdfPageSize(127d, 292.1d);         // 5 x 11.5 in
        public static PdfPageSize C => new PdfPageSize(431.8d, 558.8d);       // 17 x 22 in
        public static PdfPageSize D => new PdfPageSize(558.8d, 863.6d);       // 22 x 34 in
        public static PdfPageSize E => new PdfPageSize(863.6d, 1117.6d);      // 34 x 44 in
        public static PdfPageSize DLEnvelope => new PdfPageSize(110d, 220d);
        public static PdfPageSize C5Envelope => new PdfPageSize(162d, 229d);
        public static PdfPageSize C3Envelope => new PdfPageSize(324d, 458d);
        public static PdfPageSize C4Envelope => new PdfPageSize(229d, 324d);
        public static PdfPageSize C6Envelope => new PdfPageSize(114d, 162d);
        public static PdfPageSize C65Envelope => new PdfPageSize(114d, 229d);
        public static PdfPageSize B4Envelope => new PdfPageSize(250d, 353d);
        public static PdfPageSize B5Envelope => new PdfPageSize(176d, 250d);
        public static PdfPageSize B6Envelope => new PdfPageSize(176d, 125d);
        public static PdfPageSize ItalyEnvelope => new PdfPageSize(110d, 230d);
        public static PdfPageSize MonarchEnvelope => new PdfPageSize(98.425d, 190.5d);      // 3.875 x 7.5 in
        public static PdfPageSize Six3_4Envelope => new PdfPageSize(92.075d, 165.1d);      // 3.625 x 6.5 in
        public static PdfPageSize USStandard => new PdfPageSize(377.825d, 279.4d);     // 14.875 x 11 in
        public static PdfPageSize GermanStandard => new PdfPageSize(215.9d, 304.8d);       // 8.5 x 12 in
        public static PdfPageSize GermanLegal => new PdfPageSize(215.9d, 330.2d);       // 8.5 x 13 in
        public static PdfPageSize ISOB4 => new PdfPageSize(250d, 353d);
        public static PdfPageSize JapaneseDoublePostcard => new PdfPageSize(200d, 148d);
        public static PdfPageSize Standard9 => new PdfPageSize(228.6d, 279.4d);       // 9 x 11 in
        public static PdfPageSize Standard10 => new PdfPageSize(254d, 279.4d);         // 10 x 11 in
        public static PdfPageSize Standard15 => new PdfPageSize(381d, 279.4d);         // 15 x 11 in
        public static PdfPageSize InviteEnvelope => new PdfPageSize(220d, 220d);
        public static PdfPageSize LetterExtra => new PdfPageSize(235.585d, 304.8d);     // 9.275 x 12 in
        public static PdfPageSize LegalExtra => new PdfPageSize(235.585d, 381d);       // 9.275 x 15 in
        public static PdfPageSize TabloidExtra => new PdfPageSize(296.926d, 457.2d);     // 11.69 x 18 in
        public static PdfPageSize A4Extra => new PdfPageSize(236d, 322d);
        public static PdfPageSize LetterTransverse => new PdfPageSize(210.185d, 279.4d);     // 8.275 x 11 in
        public static PdfPageSize A4Transverse => new PdfPageSize(210d, 297d);
        public static PdfPageSize LetterExtraTransverse => new PdfPageSize(235.585d, 304.8d); // 9.275 x 12 in
        public static PdfPageSize SuperA => new PdfPageSize(227d, 356d);
        public static PdfPageSize SuperB => new PdfPageSize(305d, 487d);
        public static PdfPageSize LetterPlus => new PdfPageSize(215.9d, 322.326d);     // 8.5 x 12.69 in
        public static PdfPageSize A4Plus => new PdfPageSize(210d, 330d);
        public static PdfPageSize A5Transverse => new PdfPageSize(148d, 210d);
        public static PdfPageSize JISB5Transverse => new PdfPageSize(182d, 257d);
        public static PdfPageSize A3Extra => new PdfPageSize(322d, 445d);
        public static PdfPageSize A5Extra => new PdfPageSize(174d, 235d);
        public static PdfPageSize ISOB5 => new PdfPageSize(201d, 276d);
        public static PdfPageSize A2 => new PdfPageSize(420d, 594d);
        public static PdfPageSize A3Transverse => new PdfPageSize(297d, 420d);
        public static PdfPageSize A3ExtraTransverse => new PdfPageSize(322d, 445d);
    }
}