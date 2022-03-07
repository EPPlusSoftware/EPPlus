/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Settings for html export for ranges
    /// </summary>
    public class HtmlRangeExportSettings : HtmlExportSettings
    {
        int _headerRows=1;

        /// <summary>
        /// Number of header rows before the actual data. Default is 1.
        /// </summary>
        public int HeaderRows 
        { 
            get
            {
                return _headerRows;
            }
            set
            {
                if(value < 0 || value > ExcelPackage.MaxRows)
                {
                    throw new InvalidOperationException("Can't be negative or exceed number of allowed rows in a worksheet.");
                }
                _headerRows = value;
            }
        }
        /// <summary>
        /// If <see cref="HeaderRows"/> is 0, this collection contains the headers. 
        /// If this collection is empty the table will have no headers.
        /// </summary>
        public List<string> Headers { get; } = new List<string>();
        /// <summary>
        /// Options to exclude css elements
        /// </summary>
        public CssRangeExportSettings Css{ get; } = new CssRangeExportSettings();

        /// <summary>
        /// Reset the setting to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            Minify = true;
            HiddenRows = eHiddenState.Exclude;
            HeaderRows = 1;
            Headers.Clear();
            Accessibility.TableSettings.ResetToDefault();
            AdditionalTableClassNames.Clear();
            Culture = CultureInfo.CurrentCulture;
            Encoding = Encoding.UTF8;
            Css.ResetToDefault();
            Pictures.ResetToDefault();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(HtmlRangeExportSettings copy)
        {
            Minify = copy.Minify;
            HiddenRows = copy.HiddenRows;
            HeaderRows = copy.HeaderRows;
            Headers.Clear();
            Headers.AddRange(copy.Headers);

            Accessibility.TableSettings.Copy(copy.Accessibility.TableSettings);
            
            AdditionalTableClassNames.Clear();
            AdditionalTableClassNames.AddRange(copy.AdditionalTableClassNames);

            Culture = copy.Culture;
            Encoding = copy.Encoding;
            Css.Copy(copy.Css);
            Pictures.Copy(copy.Pictures);
        }
    }
}
