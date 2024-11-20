/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class RectLObject
    {
        /// <summary>
        /// x-coordinate of upper-left corner
        /// </summary>
        internal int Left;
        /// <summary>
        /// y-coordinate of upper-left corner
        /// </summary>
        internal int Top;

        /// <summary>
        /// x-coordinate of lower right corner
        /// </summary>
        internal int Right;
        /// <summary>
        /// y-coordinate of lower right corner
        /// </summary>
        internal int Bottom;

        internal RectLObject(BinaryReader br)
        {
            Left = br.ReadInt32();
            Top = br.ReadInt32();
            Right = br.ReadInt32();
            Bottom = br.ReadInt32();
        }

        internal RectLObject()
        {
            Left = 41;
            Top = 51;
            Right = 242;
            Bottom = 72;
        }
        internal RectLObject(int left, int top, int right, int bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(Left);
            bw.Write(Top);
            bw.Write(Right);
            bw.Write(Bottom);
        }
    }
}
