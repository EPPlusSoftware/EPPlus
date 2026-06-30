using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    public class HtmlDrawingSettings
    {
        internal HtmlDrawingSettings()
        {

        }

        //Use picture for now. Possibly re-name
        /// <summary>
        /// If how drawings should be included in the html. Default is <see cref="ePictureInclude.Exclude"/>
        /// </summary>
        public ePictureInclude Include = ePictureInclude.Exclude;

        /// <summary>
        /// Which type of drawing should be included
        /// </summary>
        public eDrawingInclude DrawTypeInclude = eDrawingInclude.None;

        /// <summary>
        /// Is absolute by default for charts
        /// </summary>
        public ePicturePosition Position = ePicturePosition.DontSet;

    }
}
