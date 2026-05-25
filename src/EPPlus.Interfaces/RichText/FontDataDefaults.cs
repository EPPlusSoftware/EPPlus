using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class FontDataDefaults : IFontData
    {
        public virtual string FamilyName { get; set; } = "Archivo Narrow";
        public virtual FontSubFamily SubFamily { get; set; } = FontSubFamily.Regular;
        public virtual float Size { get; set; } = 11f;

        public void SetFont(IFontData font)
        {
            FamilyName = font.FamilyName;
            Size = font.Size;
            SubFamily = font.SubFamily;
        }
    }
}
