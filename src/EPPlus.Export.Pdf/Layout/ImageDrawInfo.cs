/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Export.Pdf.Layout
{
    internal struct ImageDrawInfo
    {
        public double X;
        public double Y;
        public double Width;
        public double Height;
        public byte[] ImageBytes;
        public string Name;
        public string HyperlinkUri;
        public string HyperlinkTargetSheet;
        public int HyperlinkTargetRow;
        public int HyperlinkTargetColumn;
        public string HyperlinkToolTip;
    }
}
