using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;


namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    /// <summary>
    /// Interface for pdf/svg/future richtext users to unify richtext styling
    /// </summary>
    public interface IRichTextFormatSimple : IRichTextFormatBase
    {
        bool SubScript { get; set; }
        bool SuperScript { get; set; }

        /// <summary>
        /// Represents value in enum OfficeOpenXml.Style.eUnderlineType
        /// </summary>
        int UnderlineType { get; set; }
        /// <summary>
        /// Represents value in enum OfficeOpenXml.Style.eStrikeType
        /// </summary>
        int StrikeType { get; set; }

        #region potentially remove
        /// <summary>
        /// Represents OfficeOpenXml.Drawing.eTextCapsType
        /// </summary>
        int Capitalization { get; set; }
        #endregion

        Color UnderlineColor { get; set; }
        public Color FontColor { get; set; }
    }
}
