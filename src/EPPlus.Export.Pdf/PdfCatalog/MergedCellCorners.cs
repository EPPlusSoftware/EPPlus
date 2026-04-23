using System;

namespace EPPlus.Export.Pdf.PdfCatalog
{    /// <summary>
     /// Flags that describe which corners of a merged-cell boundary are owned by
     /// a particular layout object on the current page.
     ///
     /// A corner belongs to a layout object when BOTH of its edges land on the page:
     ///   TopLeft     – fromRow and fromCol are within the page range
     ///   TopRight    – fromRow and toCol   are within the page range
     ///   BottomLeft  – toRow   and fromCol are within the page range
     ///   BottomRight – toRow   and toCol   are within the page range
     ///
     /// Possible combinations:
     ///   None         – interior cell of a large merge; all four edges are
     ///                  clipped by page boundaries.
     ///   1 flag       – only one corner of the merge lands on this page.
     ///   2 flags      – merge is exactly 1 row tall  → TopLeft|TopRight  or BottomLeft|BottomRight
     ///                  merge is exactly 1 col wide  → TopLeft|BottomLeft or TopRight|BottomRight
     ///   All (4 flags)– non-merged cell, or merged cell fully contained on this page.
     /// </summary>
    [Flags]
    internal enum MergedCellCorners
    {
        None = 0,
        TopLeft = 1 << 0,   // 1
        TopRight = 1 << 1,   // 2
        BottomLeft = 1 << 2,   // 4
        BottomRight = 1 << 3,   // 8
        All = TopLeft | TopRight | BottomLeft | BottomRight
    }
}
