using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;


namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Interface for pdf/svg/future richtext users to unify richtext styling
    /// </summary>
    public interface IRichTextInfoBase : IFontData
    {
        /// <summary>
        /// This MUST interact with font data subfamily
        /// </summary>
        bool Italic { get; set; }
        /// <summary>
        /// This MUST interact with font data subfamily
        /// </summary>
        bool Bold { get; set; }
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

        Color UnderlineColor { get; set; }
        public Color FontColor { get; set; }

        //TODO: ColorSettings independent from Epplus
    }
}
