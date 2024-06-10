using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Chart.DataLabling
{
    internal class CTDataLabel : GroupDataLabel
    {
        string[] nodeOrder;

        internal CTDataLabel()
        {
            nodeOrder = new string[GroupDLbl.Length + 3];

            nodeOrder[0] = "idx";
            //Note: this delete and group is one "choice" node.
            nodeOrder[1] = "delete";

            ////Alternate:
            //Array.Copy(EG_DlblShared, _grpLbl, EG_DlblShared.Length);

            for (int i = 0; i < GroupDLbl.Length; i++)
            {
                nodeOrder[i + 2] = GroupDLbl[i];
            }

            nodeOrder[nodeOrder.Length - 1] = "extLst";

            NodeOrder = nodeOrder;
        }

        internal string[] NodeOrder { get; private set; }
    }
}
