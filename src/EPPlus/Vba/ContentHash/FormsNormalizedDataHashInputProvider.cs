using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal class FormsNormalizedDataHashInputProvider : ContentHashInputProvider
    {
        public FormsNormalizedDataHashInputProvider(ExcelVbaProject project) : base(project)
        {
        }

        protected override void CreateHashInputInternal(MemoryStream ms)
        {
            //MS-OVBA 2.4.2.2
            BinaryWriter bw = new BinaryWriter(ms);
            
            // Todo: write input to stream...
        }
    }
}
