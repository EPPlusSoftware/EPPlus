/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.RichData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;

namespace OfficeOpenXml.CellPictures
{
    /// <summary>
    /// Represents an in-cell picture
    /// </summary>
    internal class ExcelCellPicture
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelCellPicture()
        {
            
        }

        /// <summary>
        /// Internal uri in the workbook of the image.
        /// </summary>
        public Uri ImageUri
        {
            get; set;
        }

        /// <summary>
        /// Alt text of the image
        /// </summary>
        public string AltText
        {
            get; set;
        }

        /// <summary>
        /// Indicates the calculation context in which this image was created.
        /// </summary>
        internal CalcOrigins CalcOrigin { get; set; }

        /// <summary>
        /// Address of the cell picture
        /// </summary>
        public ExcelAddress CellAddress { get; set;  }

       
    }
}
