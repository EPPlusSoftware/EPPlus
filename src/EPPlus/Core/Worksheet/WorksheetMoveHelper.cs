/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class MoveSheetXmlNode
    {       
        internal static void RearrangeWorksheets(ExcelWorksheets worksheets, string sourceWorksheetName, string targetWorksheetName, bool before)
        {
            var sourceWorksheet = worksheets[sourceWorksheetName];
            var targetWorksheet=worksheets[targetWorksheetName];
            if (sourceWorksheet == null)
            {
                throw new ArgumentException($"Could not find source worksheet {sourceWorksheet} to move.");
            }
            if (targetWorksheet == null)
            {
                throw new ArgumentException($"Could not find target worksheet {targetWorksheet} to move.");
            }
            RearrangeWorksheets(sourceWorksheet.Workbook.Worksheets, sourceWorksheet.PositionId, targetWorksheet.PositionId, before);
        }

        internal static void RearrangeWorksheets(ExcelWorksheets worksheets, int sourcePositionId, int targetPositionId, bool before)
        {
            if (sourcePositionId == targetPositionId) return;

            lock (worksheets)
            {
                var sourceSheet = worksheets[sourcePositionId];
                var targetSheet = worksheets[targetPositionId];

                var package = targetSheet._package;

                // Preserve ActiveTab after rearrange - save active tab if set
                var activeSheet = default(ExcelWorksheet);
                var activeSheetIx = int.MinValue;
                if (package.Workbook.View != null
                    && package.Workbook.View.ActiveTab >= 0
                    && package.Workbook.View.ActiveTab < worksheets.Count)
                {
                    activeSheetIx = package.Workbook.View.ActiveTab;
                    activeSheet = package.Workbook.Worksheets[activeSheetIx];
                }


                var index = package._worksheetAdd;

                worksheets._worksheets.Move(sourcePositionId - index, targetPositionId - index, before);

                var from = Math.Min(sourcePositionId, targetPositionId);
                var to = Math.Max(sourcePositionId, targetPositionId);
                for (int i = from; i <= to; i++)
                {
                    worksheets[i].PositionId = i;
                }

                MoveTargetXml(worksheets, sourceSheet, targetSheet, before);

                // Preserve ActiveTab after rearrange
                if (activeSheet != null)
                {
                    package.Workbook.View.ActiveTab = activeSheet.PositionId - package._worksheetAdd;
                }
            }
        }

        private static void MoveTargetXml(ExcelWorksheets worksheets, ExcelWorksheet sourceWs, ExcelWorksheet targetWs, bool before)
        {
            var sourceNode = worksheets.TopNode.SelectSingleNode($"d:sheet[@sheetId = '{sourceWs.SheetId}']", worksheets.NameSpaceManager);
            var targetNode = worksheets.TopNode.SelectSingleNode($"d:sheet[@sheetId = '{targetWs.SheetId}']", worksheets.NameSpaceManager);
            if (sourceNode == null || targetNode == null)
            {
                throw new InvalidOperationException("Invalid Workbook Xml. Can't find worksheet in workbook list.");
            }
            if (before)
            {
                worksheets.TopNode.InsertBefore(sourceNode, targetNode);
            }
            else
            {
                worksheets.TopNode.InsertAfter(sourceNode, targetNode);
            }
        }
    }
}
