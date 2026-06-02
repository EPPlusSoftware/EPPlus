using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    /// <summary>
    /// The most basic rich text format
    /// The only properties that belong in this class are those that are absolutely neccesary for Measuring the text correctly
    /// </summary>
    public class RichTextFormatBase : FontFormatBase, IRichTextFormatBase
    {
        internal RichTextFormatBase() 
        {
            Italic = false;
            Bold = false;
        }

        public RichTextFormatBase(string text, string fontFamily, float size, bool bold = false, bool italic = false)
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
