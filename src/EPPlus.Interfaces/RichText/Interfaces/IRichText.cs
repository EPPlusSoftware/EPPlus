using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    public interface IRichText
    {
        string Text { get; set; }

        IRichTextInfoBase RichTextOptions { get; set; }
    }
}
