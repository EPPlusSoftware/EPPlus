/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.EventArguments
{
    internal class WorksheetValueMetadataReadEventArgs : EventArgs
    {
        public WorksheetValueMetadataReadEventArgs(int worksheetIx, int row, int col, uint vm)
        {
            WorksheetIx = worksheetIx;
            Row = row;
            Col = col;
            Vm = vm;
        }
        public int WorksheetIx { get; }
        public int Row { get; }
        public int Col { get; }
        public uint Vm { get; }
    }
}
