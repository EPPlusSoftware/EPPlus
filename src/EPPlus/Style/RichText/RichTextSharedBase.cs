using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// Common Base class for ALL Epplus RichText
    /// </summary>
    public class RichTextSharedBase : FontDataBasic, IRichTextSharedBase
    {
        internal RichTextSharedBase()
        {
            Italic = false;
            Bold = false;
        }

        /// <summary>
        /// The text string
        /// </summary>
        public virtual string Text { get; set; }

        /// <summary>
        /// FontItalic text
        /// </summary>
        public virtual bool Italic
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
        /// FontBold text
        /// </summary>
        public virtual bool Bold
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
