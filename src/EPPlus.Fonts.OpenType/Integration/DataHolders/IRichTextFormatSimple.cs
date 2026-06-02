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
    /// TODO: Move this to interfaces. Only here in order to not break existing references in PDF
    /// Interface for pdf/svg/future richtext users to unify richtext styling
    /// </summary>
    public interface IRichTextFormatSimple : IRichTextFormatEssential
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
