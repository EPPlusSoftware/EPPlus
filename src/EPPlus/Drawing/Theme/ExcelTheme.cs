/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// Defines a Theme within the package
    /// </summary>
    public class ExcelTheme : ExcelThemeBase
    {
        internal ExcelWorkbook _wb;
        /// <summary>
        /// The name of the theme
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }

        internal ExcelTheme(ExcelWorkbook workbook, ZipPackageRelationship rel)
            : base(workbook._package,workbook.NameSpaceManager, rel, "a:themeElements/")
        {
            _wb = workbook;
        }
    }
}
