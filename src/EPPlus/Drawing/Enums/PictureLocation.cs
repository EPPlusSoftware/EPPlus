using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Options for how to link a picture
    /// </summary>
    [Flags]
    public enum PictureLocation
    {
        /// <summary>
        /// 
        /// </summary>
        None = 0x00,
        /// <summary>
        /// Copy and Embed the image within the workbook
        /// </summary>
        Embed = 0x01,
        /// <summary>
        /// Collect the image from the link
        /// </summary>
        Link = 0x02,
        /// <summary>
        /// Copy and Embed the image and add a link
        /// </summary>
        LinkAndEmbed = Embed | Link
    }
}
