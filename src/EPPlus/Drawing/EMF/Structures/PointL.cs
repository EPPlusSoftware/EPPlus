using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class PointL
    {
        internal int PointX;
        internal int PointY;

        internal PointL(BinaryReader br)
        {
            PointX = br.ReadInt32();
            PointY = br.ReadInt32();
        }

        internal PointL(int x, int y)
        {
            PointX = x;
            PointY = y;
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(PointX);
            bw.Write(PointY);
        }
    }
}
