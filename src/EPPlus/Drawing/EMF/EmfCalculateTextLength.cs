using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EmfCalculateTextLength
    {
        internal struct RowData
        {
            internal int Length;
            internal string Text;
            internal int PosX;
            internal int PosY;
        }

        internal List<EMR_RECORD> TextRecords = new List<EMR_RECORD>();

        //For the EMF image created by OLE objects, the text has the following properties:
        //We can fit 32 of the smalest character, l
        //This means that max length = 96. We will give some space and set length = 90, which fits 30 characters.
        //This gives us a range from 0 - 47, where 47 is the center.
        //We can fit 11 of the widest character, O.
        //This means that max length = 99. We will give it less space and set length = 90, which fits 10 characters.
        //This gives us a range from 0 - 44, where 44 is the center.
        //This gives us a conversion range of 3 - 96 character lengths to range 45 - 1 start x.

        private int minWidth = 90;
        private int maxWidth = 96;
        private int characterWidthRangeLow = 3;
        private int characterWidthRangeHigh = 96;
        private int textStartLocationLow = 45;
        private int textStartLocationHigh = 1;
        private int maxRows = 3;
        private int rowIndex = 0;
        private int rowstartY = 36;
        private int rowYIncrement = 12;

        internal EmfCalculateTextLength(string text)
        {
            List<RowData> rows = new List<RowData>();
            int rowCharactersWidth = 0;
            int rowCharacterNumber = 0;
            int rowStartIndex = 0;

            //Get substrings and row lengths
            for (int i = 0; i < text.Length; i++)
            {
                rowCharactersWidth += EMR_EXTTEXTOUTW.GetSpacingForChar(text[i]);
                rowCharacterNumber++;
                if(rowCharactersWidth >= minWidth || i == text.Length-1)
                {
                    //If we exceed max length, we don't include the last character.
                    if(rowCharactersWidth >= maxWidth)
                    {
                        i--;
                    }
                    RowData rowData = new RowData();
                    rowData.Text = text.Substring(rowStartIndex, rowCharacterNumber);
                    rowData.Length = rowCharactersWidth;
                    rowData.PosX = ConvertRange(characterWidthRangeLow, characterWidthRangeHigh, textStartLocationLow, textStartLocationHigh, rowCharactersWidth);
                    rowData.PosY = rowstartY;
                    rows.Add(rowData);
                    rowstartY += rowYIncrement;
                    rowStartIndex = i + 1;
                    rowCharactersWidth = 0;
                    rowCharacterNumber = 0;
                    rowIndex++;
                    if (rowIndex > maxRows-1)
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

        private static int ConvertRange(int OriginalRangeStart, int OriginalRangeEnd, int NewRangeStart, int NewRangeEnd, int OriginalValue)
        {
            double scaling = (double)(NewRangeEnd - NewRangeStart) / (OriginalRangeEnd - OriginalRangeStart);
            return (int)(NewRangeStart + ((OriginalValue - OriginalRangeStart) * scaling));
        }
    }
}
