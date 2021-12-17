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
using System.Globalization;
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
                Culture = CultureInfo.CurrentCulture
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

        private List<string> _additionalTableClasses = new List<string>(20);
        

        /// <summary>
        /// Use this property to set additional class names that will be set on the exported html-table. Max 20 additional classes can be added
        /// </summary>
        public IList<string> AdditionalTableClassNames
        {
            get
            {
                return _additionalTableClasses;
            }
        }

        public CultureInfo Culture { get; set; }

        /// <summary>
        /// If true data-* attributes will be rendered
        /// </summary>
        public bool RenderDataAttributes { get; set; }
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
        internal Dictionary<string, string> AdditionalCssElements
        {
            get;
            private set;
        } = new Dictionary<string, string>()
            {
                { "border-spacing", "0" },
                { "border-collapse", "collapse" },
                { "word-wrap", "break-word"},
                { "white-space", "nowrap"}
            };

        /// <summary>
        /// Include Css for the current table style
        /// </summary>
        public bool IncludeTableStyles { get; set; } = true;

        public float Indent { get; set; } = 2;
        public string IndentUnit { get; set; } = "em";
        public CssExcludeStyle Exclude
        {
            get;
        } = new CssExcludeStyle();
    }
    [Flags]
    public enum eFontExclude
    {
        Name = 0x01,
        Size = 0x02,
        Color = 0x04,
        Bold = 0x08,
        Italic = 0x10,
        Strike = 0x20,
        Underline = 0x40,
    }
    [Flags]
    public enum eBorderExclude
    {
        Top = 0x01,
        Bottom = 0x02,
        Left = 0x04,
        Right = 0x08
    }

    public class CssExcludeStyle
    {
        public CssExclude TableStyle { get; } = new CssExclude();
        public CssExclude CellStyle { get; } = new CssExclude();
    }
    public class CssExclude
    {
        public eFontExclude Font { get; set; }
        public eBorderExclude Border { get; set; }
        public bool Fill { get; set; }
        public bool VerticalAlignment { get; set; }
        public bool HorizontalAlignment { get; set; }
        public bool WrapText { get; set; }
        public bool TextRotation { get; set; }
        public bool Indent { get; set; }
    }

    public abstract class HtmlTableExportBaseOptions
    {
        /// <summary>
        /// If set to true the rendered html will be formatted with indents and linebreaks.
        /// </summary>
        public bool Minify { get; set; } = true;
        /// <summary>
        /// Include Css for cell styling.
        /// </summary>
        public bool IncludeCellStyles { get; set; } = true;
    }
}
