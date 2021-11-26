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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Settings for the 
    /// </summary>
    public class HtmlTableExportSettings
    {
        /// <summary>
        /// If set to true the rendered html will be formatted with indents and linebreaks.
        /// </summary>
        public bool Minify { get; set; } = true;
        /// <summary>
        /// If true hidden rows will be included. 
        /// </summary>
        public bool IncludeHiddenRows { get; set; } = false;
        /// <summary>
        /// Settings for usage of accessibility (aria, role) attributes of the table
        /// </summary>
        public AccessibilitySettings Accessibility
        {
            get; private set;
        } = new AccessibilitySettings();
        /// <summary>
        /// If set to true classes that identifies Excel table styling will be included in the html. Default value is true.
        /// </summary>
        public bool IncludeDefaultClasses { get; set; } = true;
        /// <summary>
        /// The html id attribute for the exported table. The id attribute is only added to the table if this property is not null or empty.
        /// </summary>
        public string TableId { get; set; }

        /// <summary>
        /// Use this property to set additional class names that will be set on the exported html-table.
        /// </summary>
        public IList<string> AdditionalTableClassNames
        {
            get;
            private set;
        } = new List<string>();

        /// <summary>
        /// The culture used when formatting the cell output.
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;
        public Encoding Encoding { get; set; } = Encoding.UTF8;
        /// <summary>
        /// If true data-* attributes will be rendered
        /// </summary>
        public bool RenderDataAttributes { get; set; } = true;

        public CssTableExportSettings Css { get; } = new CssTableExportSettings();

        public void ResetToDefault()
        {
            Minify = true;
            IncludeHiddenRows = false;
            Accessibility.TableSettings.ResetToDefault();
            IncludeDefaultClasses = true;
            TableId = "";
            AdditionalTableClassNames=new List<string>();
            Culture = CultureInfo.CurrentCulture;
            Encoding = Encoding.UTF8;
            RenderDataAttributes = true;
            Css.ResetToDefault();
        }
        public void Copy(HtmlTableExportSettings copy)
        {
            Minify = copy.Minify;
            IncludeHiddenRows = copy.IncludeHiddenRows;
            Accessibility.TableSettings.Copy(copy.Accessibility.TableSettings);
            IncludeDefaultClasses = copy.IncludeDefaultClasses;
            TableId = copy.TableId;
            AdditionalTableClassNames = copy.AdditionalTableClassNames;
            Culture = copy.Culture;
            Encoding = copy.Encoding;
            RenderDataAttributes = copy.RenderDataAttributes;
            Css.Copy(copy.Css);
        }
        public void Configure(Action<HtmlTableExportSettings> settings)
        {
            settings.Invoke(this);
        }
    }
    public class CssTableExportSettings
    {
        internal CssTableExportSettings()
        {
            ResetToDefault();
        }
        /// <summary>
        /// Include Css for the current table style
        /// </summary>
        public bool IncludeTableStyles { get; set; } = true;
        /// <summary>
        /// Include Css for cell styling.
        /// </summary>
        public bool IncludeCellStyles { get; set; } = true;
        /// <summary>
        /// Css elements added to the table.
        /// </summary>
        internal Dictionary<string, string> AdditionalCssElements
        {
            get;
            private set;
        }
        /// <summary>
        /// The value used in the stylesheet for an indentation in a cell
        /// </summary>
        public float IndentValue { get; set; } = 2;
        /// <summary>
        /// The unit used in the stylesheet for an indentation in a cell
        /// </summary>
        public string IndentUnit { get; set; } = "em";
        /// <summary>
        /// Exclude flags for styles
        /// </summary>
        public CssExcludeStyle Exclude
        {
            get;
        } = new CssExcludeStyle();

        public void ResetToDefault()
        {
            IncludeTableStyles = true;
            IncludeCellStyles = true;
            AdditionalCssElements = new Dictionary<string, string>()
            {
                { "border-spacing", "0" },
                { "border-collapse", "collapse" },
                { "word-wrap", "break-word"},
                { "white-space", "nowrap"}
            };
            IndentValue = 2;
            IndentUnit = "em";

            Exclude.TableStyle.ResetToDefault();
            Exclude.CellStyle.ResetToDefault();
        }

        internal void Copy(CssTableExportSettings copy)
        {
            IncludeTableStyles = copy.IncludeTableStyles;
            IncludeCellStyles = copy.IncludeTableStyles;
            AdditionalCssElements = copy.AdditionalCssElements;
            IndentValue = copy.IndentValue;
            IndentUnit = copy.IndentUnit;

            Exclude.TableStyle.Copy(copy.Exclude.TableStyle);
            Exclude.CellStyle.Copy(copy.Exclude.CellStyle);
        }
    }
    [Flags]
    public enum eFontExclude
    {
        All = 0x4F,
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
        All = 0x0F,
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

        public void ResetToDefault()
        {
            Font = 0;
            Border = 0;
            Fill = false;
            VerticalAlignment = false;
            HorizontalAlignment = false;
            WrapText = false;
            TextRotation = false;
            Indent = false;
        }
        public void Copy(CssExclude copy)
        {
            Font = copy.Font;
            Border = copy.Border;
            Fill = copy.Fill;
            VerticalAlignment = copy.VerticalAlignment;
            HorizontalAlignment = copy.HorizontalAlignment;
            WrapText = copy.WrapText;
            TextRotation = copy.TextRotation;
            Indent = copy.Indent;
        }        
    }
}
