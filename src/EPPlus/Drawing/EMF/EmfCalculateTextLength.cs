using OfficeOpenXml.Drawing.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EmfCalculateTextLength
    {
        struct RowData
        {
            internal int Length;
            internal string Text;
            internal int PosX;
            internal int PosY;
        }

        internal List<EMR_RECORD> TextRecords = new List<EMR_RECORD>();

        int minLength = 90;
        int maxLength = 96;
        int maxRows = 3;

        public EmfCalculateTextLength(string text)
        {
            List<RowData> rows = new List<RowData>();
            int rowLen=0;
            int rowStartIndex=0;
            int rowIndex=1;
            int rowstartY = 36;
            //Get substrings and row lengths
            for (int i = 0; i < text.Length; i++)
            {
                rowLen += EMR_EXTTEXTOUTW.GetSpacingForChar(text[i]);
                if(rowLen >= minLength)
                {
                    //If we exceed max length, we don't include the last character.
                    if(rowLen >= maxLength)
                    {
                        i--;
                    }
                    RowData rowData = new RowData();
                    rowData.Text = text.Substring(rowStartIndex, i);
                    rowData.Length = rowLen;
                    rowData.PosX = ConvertRange(3, 96, 45, 0, rowLen);
                    rowData.PosY = rowstartY;
                    rows.Add(rowData);
                    rowstartY += 12;
                    rowStartIndex = i+1;
                    rowIndex++;
                    if(rowIndex > maxRows)
                    {
                        break;
                    }
                }
            }
            //Create TextRecords
            for (int i = 0; i < rows.Count; i++)
            {
                LogFontExDv elw = new LogFontExDv();
                elw.Height = unchecked((int)4294967285);
                elw.Width = 0;
                elw.Escapement = 0;
                elw.Orientation = 0;
                elw.Weight = 400;
                elw.Italic = 0;
                elw.Underline = 0;
                elw.StrikeOut = 0;
                elw.Set = 0;
                elw.OutPrecision = 0;
                elw.ClipPrecision = 0;
                elw.Quality = 0;
                elw.PitchAndFamily = 32;
                elw.FaceName = "Tahoma";
                elw.FullName = string.Empty;
                elw.Style = string.Empty;
                elw.Script = string.Empty;
                elw.dv.Signature = 134248036;
                elw.dv.NumAxes = 0;
                elw.dv.Values = new uint[0] { };
                EMR_EXTCREATEFONTINDIRECTW font = new EMR_EXTCREATEFONTINDIRECTW(elw);
                TextRecords.Add(font);
                EMR_SELECTOBJECT selObj1 = new EMR_SELECTOBJECT(2);
                TextRecords.Add(selObj1);
                EMR_SETBKMODE bkMode = new EMR_SETBKMODE(1);
                TextRecords.Add(bkMode);
                EMR_EXTTEXTOUTW textw = new EMR_EXTTEXTOUTW(rows[i].Text, rows[i].PosX, rows[i].PosY);
                TextRecords.Add(textw);
                EMR_SELECTOBJECT selObj2 = new EMR_SELECTOBJECT(2147483661);
                TextRecords.Add(selObj2);
                EMR_DELETEOBJECT delObj = new EMR_DELETEOBJECT(4294967295);
                TextRecords.Add(delObj);
            }
        }

        private static int ConvertRange(int originalStart, int originalEnd, int newStart, int newEnd, int value)
        {
            double scale = (double)(newEnd - newStart) / (originalEnd - originalStart);
            return (int)(newStart + ((value - originalStart) * scale));
        }

    }
}
