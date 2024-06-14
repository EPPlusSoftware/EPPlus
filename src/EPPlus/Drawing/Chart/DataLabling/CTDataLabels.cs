using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Chart.DataLabling
{
    internal class CTDataLabels : GroupDataLabels
    {
        string[] _nodeOrder;

        internal CTDataLabels()
        {
            _nodeOrder = new string[GroupDLbls.Length + 3];
            _nodeOrder[0] = "dLbl";
            _nodeOrder[1] = "delete";

            ////Alternative 1:
            //Array.Copy(EG_DlblShared, _grpLbls, EG_DlblShared.Length);

            ////Alternative 2:
            for (int i = 0; i < GroupDLbls.Length; i++)
            {
                _nodeOrder[i + 2] = GroupDLbls[i];
            }

            _nodeOrder[_nodeOrder.Length - 1] = "extLst";

            NodeOrder = _nodeOrder;
        }

        internal string[] NodeOrder { get; private set; }
    }
}
