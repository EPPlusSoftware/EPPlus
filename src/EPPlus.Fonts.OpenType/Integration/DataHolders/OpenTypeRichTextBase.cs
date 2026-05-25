using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    public class OpenTypeRichTextBase : OpenTypeFontInfoBase, IRichTextFormatBase
    {
        public OpenTypeRichTextBase(string text, string fontFamily, float size, bool bold, bool italic)
        {
            Text = text;
            Family = fontFamily;
            Size = size;
            Bold = bold;
            Italic = italic;
        }

        public string Text { get; set; }

        /// <summary>
        /// Any inheriting class MUST do this too
        /// </summary>
        public bool Italic
        {
            get { return (SubFamily & FontSubFamily.Italic) == FontSubFamily.Italic; }
            set
            {
                if (value)
                {
                    //Set Flag
                    SubFamily = SubFamily | FontSubFamily.Italic;

                }
                else
                {
                    //Unset Flag
                    SubFamily &= ~FontSubFamily.Italic;
                }
            }
        }

        /// <summary>
        /// Any inheriting class MUST do this too
        /// </summary>
        public bool Bold
        {
            get { return (SubFamily & FontSubFamily.Bold) == FontSubFamily.Bold; }
            set
            {
                if (value)
                {
                    //Set flag
                    SubFamily = SubFamily | FontSubFamily.Bold;

                }
                else
                {
                    //Unset flag
                    SubFamily &= ~FontSubFamily.Bold;
                }
            }
        }
    }
}
