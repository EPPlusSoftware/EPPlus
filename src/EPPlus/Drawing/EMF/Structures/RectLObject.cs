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

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(Left);
            bw.Write(Top);
            bw.Write(Right);
            bw.Write(Bottom);
        }
    }
}
