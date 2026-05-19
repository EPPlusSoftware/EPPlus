using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.Drawing.RichText
{
    /// <summary>
    /// For ln Outline
    /// </summary>
    public enum TextCapsType
    {
        /// <summary>
        /// Apply all caps on the text. All lower case letters are converted to upper case, but stored without change.
        /// </summary>
        All,
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// Apply small caps to the text. Letters are converted to lower case.
        /// </summary>
        Small
    }
}
