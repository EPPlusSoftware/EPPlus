using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.RichText;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal interface IRichTextInfoDrawing : IRichTextInfoBase
    {
        DrawingStrikeType DrawingStrike{ get; set; } /*{ get { return (DrawingStrikeType)StrikeType; } set { StrikeType = (int)value; } }*/
        DrawingUnderlineStyle UnderlineStyle { get; set; }
        DrawingTextCapsType Capitalization { get; set; }
    }
}
