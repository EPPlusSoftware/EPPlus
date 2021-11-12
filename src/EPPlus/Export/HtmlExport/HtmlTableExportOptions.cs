/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// 
    /// </summary>
    public class HtmlTableExportOptions : HtmlTableExportBaseOptions
    {
        internal HtmlTableExportOptions() : base() { }

        public static HtmlTableExportOptions Create()
        {
            var defaultOptions = new HtmlTableExportOptions
            {
                IncludeDefaultClasses = true,
                Minify = true,
            };
            return defaultOptions;            
        }
        internal static HtmlTableExportOptions Default
        {
            get { return Create(); }
        }

        /// <summary>
        /// If set to true classes that identifies Excel table styling will be included in the html. Default value is true.
        /// </summary>
        public bool IncludeDefaultClasses { get; set; }

        /// <summary>
        /// The html id attribute for the exported table. The id attribute is only added to the table if this property is not null or empty.
        /// </summary>
        public string TableId { get; set; }

        /// <summary>
        /// If true data-* attributes will be rendered
        /// </summary>
        public bool RenderDataAttributes { get; set; }
        public bool IncludeCss { get; set; }
    }
    public class CssTableExportOptions : HtmlTableExportBaseOptions
    {
        public static CssTableExportOptions Create()
        {
            var defaultOptions = new CssTableExportOptions
            {
                Minify = true
            };
            return defaultOptions;
        }
        internal static CssTableExportOptions Default
        {
            get { return Create(); }
        }
    }
    public abstract class HtmlTableExportBaseOptions
    {
        /// <summary>
        /// If set to true the rendered html will be formatted with indents and linebreaks.
        /// </summary>
        public bool Minify { get; set; } = true;
    }
}
