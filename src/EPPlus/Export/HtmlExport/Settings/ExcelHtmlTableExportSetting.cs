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
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Settings
{
    /// <summary>
    /// Settings for html export for tables
    /// </summary>
    public class HtmlTableExportSettings : HtmlExportSettings
    {
        /// <summary>
        /// Css export settings.
        /// </summary>
        public CssTableExportSettings Css { get; } = new CssTableExportSettings();

        /// <summary>
        /// Reset the settings to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            Minify = true;
            HiddenRows = eHiddenState.Exclude;
            Accessibility.TableSettings.ResetToDefault();
            IncludeCssClassNames = true;
            TableId = "";
            AdditionalTableClassNames.Clear();
            Culture = CultureInfo.CurrentCulture;
            Encoding = Encoding.UTF8;
            RenderDataAttributes = true;
            Css.ResetToDefault();
            Pictures.ResetToDefault();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(HtmlTableExportSettings copy)
        {
            Minify = copy.Minify;
            HiddenRows = copy.HiddenRows;
            Accessibility.TableSettings.Copy(copy.Accessibility.TableSettings);
            IncludeCssClassNames = copy.IncludeCssClassNames;
            TableId = copy.TableId;
            AdditionalTableClassNames = copy.AdditionalTableClassNames;
            Culture = copy.Culture;
            Encoding = copy.Encoding;
            RenderDataAttributes = copy.RenderDataAttributes;
            Css.Copy(copy.Css);
            Pictures.Copy(copy.Pictures);
        }
        /// <summary>
        /// Configure the settings.
        /// </summary>
        /// <param name="settings"></param>
        public void Configure(Action<HtmlTableExportSettings> settings)
        {
            settings.Invoke(this);
        }
    }
}
