using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.String
{
    internal class UTF8Upper : UTF8Encoding
    {
        internal UTF8Upper(bool AddBOM = false) : base(AddBOM)
        {
        }
        public override string BodyName => base.BodyName?.ToUpper();
        public override string HeaderName => base.HeaderName?.ToUpper();
        public override string WebName => base.WebName?.ToUpper();
    }
}
