using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Chart.DataLabling
{
    internal class GroupDataLabels : DataLabelConstNodeOrderBase
    {
        string[] _grpLbls = new string[EG_DlblShared.Length + 2];

        internal GroupDataLabels()
        {
            _grpLbls[_grpLbls.Length - 2] = "showLeaderLines";
            _grpLbls[_grpLbls.Length - 1] = "leaderLines";

            ////Alternative 1:
            //Array.Copy(EG_DlblShared, _grpLbls, EG_DlblShared.Length);

            ////Alternative 2:
            for (int i = 0; i < EG_DlblShared.Length; i++)
            {
                _grpLbls[i] = EG_DlblShared[i];
            }

            GroupDLbls = _grpLbls;
        }

        internal static string[] GroupDLbls { get; private set; }
    }
}
