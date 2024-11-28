using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF.PlusStructure
{
    internal class EmfPlusGraphicsVersionObject
    {
        internal byte[] bytes;
        internal EmfPlusGraphicsVersionObject(BinaryReader br)
        {
            //Remember: last becomes first because endian

            bytes = br.ReadBytes(4);
            //TODO:
            //Somehow breakout the two variables
            //The first 20 *Bits* is MetafileSignature so 2 and 1/2 bytes
            //The next 12  *Bits* is GraphicsVersion 1/2 and 1 bytes
        }
    }
}
