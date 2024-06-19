using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Chart.DataLabling
{
    internal class GroupDataLabel : DataLabelConstNodeOrderBase
    {
        string[] _grpLbl = new string[EG_DlblShared.Length + 2];

        internal GroupDataLabel()
        {
            _grpLbl[0] = "layout";
            _grpLbl[1] = "tx";

            ////Alternate:
            //Array.Copy(EG_DlblShared, _grpLbl, EG_DlblShared.Length);

            for (int i = 0; i < EG_DlblShared.Length; i++)
            {
                _grpLbl[i + 2] = EG_DlblShared[i];
            }

            GroupDLbl = _grpLbl;
        }

        internal static string[] GroupDLbl { get; private set; }
    }
}
