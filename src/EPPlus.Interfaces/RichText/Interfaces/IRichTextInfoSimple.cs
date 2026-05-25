using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;


namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Simple rich-text data like e.g. the data in Excel cells
    /// </summary>
    public interface IRichTextInfoSimple : IRichTextInfoEssential
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

        Color UnderlineColor { get; set; }
        public Color FontColor { get; set; }

        //TODO: ColorSettings independent from Epplus
    }
}
