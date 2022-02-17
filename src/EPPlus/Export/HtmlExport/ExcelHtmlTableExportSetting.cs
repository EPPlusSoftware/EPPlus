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
    public enum eHtmlGeneralAlignmentHandling
    {
        /// <summary>
        /// Dont set any alignment when alignment is set to general
        /// </summary>
        DontSet,
        /// <summary>
        /// If the column data type is numeric or date, alignment will be right otherwise left.
        /// </summary>
        ColumnDataType,
        /// <summary>
        /// If the cell value data type is numeric or date, alignment will be right otherwise left.
        /// </summary>
        CellDataType
    }
    public enum eHiddenState
    {
        Exclude,
        IncludeButHide,
        Include
    }
    public abstract class HtmlExportSettings
    {
        /// <summary>
        /// The html id attribute for the exported table. The id attribute is only added to the table if this property is not null or empty.
        /// </summary>
        public string TableId { get; set; }
        /// <summary>
        /// If set to true the rendered html will be formatted with indents and linebreaks.
        /// </summary>
        public bool Minify { get; set; } = true;
        /// <summary>
        /// How hidden rows will be handled. Default is <see cref="eHiddenState.Exclude"/> 
        /// </summary>
        public eHiddenState HiddenRows { get; set; } = eHiddenState.Exclude;
        /// <summary>
        /// How to set the alignment for a cell if it's alignment is set to General.
        /// </summary>
        public eHtmlGeneralAlignmentHandling HorizontalAlignmentWhenGeneral { get; set; } = eHtmlGeneralAlignmentHandling.CellDataType;
        /// <summary>
        /// Settings for usage of accessibility (aria, role) attributes of the table
        /// </summary>
        public AccessibilitySettings Accessibility
        {
            get; private set;
        } = new AccessibilitySettings();
        /// <summary>
        /// Use this property to set additional class names that will be set on the exported html-table.
        /// </summary>
        public List<string> AdditionalTableClassNames
        {
            get;
            protected internal set;
        } = new List<string>();

        /// <summary>
        /// The culture used when formatting the cell output.
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;
        /// <summary>
        /// Encoding for the output
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;
        /// <summary>
        /// Set the column width for columns in the table via the columngroup/col element.
        /// Columns with the default width will have the default column width class set, ({Settings.StyleClassPrefix}dcw). 
        /// Columns with custom column width will have the width set directly via the style attribute.
        /// </summary>
        public bool SetColumnWidth { get; set; } = false;
        /// <summary>
        /// Set the row height for rows in the table.
        /// Rows with the default height will have the default row height class set, ({Settings.StyleClassPrefix}drh). 
        /// Rows with custom row height will have the height set directly via the style attribute.
        /// </summary>
        public bool SetRowHeight { get; set; } = false;
        /// <summary>
        /// Prefix for style classes added by EPPlus. 
        /// <list type="table">
        /// <listheader><term>type</term><term>Class name</term></listheader>
        /// <item><term>Cell styles</term><term>{prefix}s{index}</term></item>
        /// <item><term>Alignment Left </term><term>{prefix}al}</term></item>
        /// <item><term>Alignment Right</term><term>{prefix}ar}</term></item>
        /// <item><term>Default column width</term><term>{prefix}dcw}</term></item>
        /// <item><term>Default row height</term><term>{prefix}drh}</term></item>
        /// </list>
        /// </summary>
        public string StyleClassPrefix { get; set; } = "epp-";
    }

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
        }
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
        }
    }
    /// <summary>
    /// Settings for html export for tables
    /// </summary>
    public class HtmlTableExportSettings : HtmlExportSettings
    {
        /// <summary>
        /// If set to true classes that identifies Excel table styling will be included in the html. Default value is true.
        /// </summary>
        public bool IncludeDefaultClasses { get; set; } = true;
        /// <summary>
        /// If true data-* attributes will be rendered
        /// </summary>
        public bool RenderDataAttributes { get; set; } = true;
        /// <summary>
        /// If true, data types are renedered on the header objects.
        /// </summary>
        public bool RenderDataTypes { get; set; } = true;


        public CssTableExportSettings Css { get; } = new CssTableExportSettings();

        public void ResetToDefault()
        {
            Minify = true;
            HiddenRows = eHiddenState.Exclude;
            Accessibility.TableSettings.ResetToDefault();
            IncludeDefaultClasses = true;
            TableId = "";
            AdditionalTableClassNames.Clear();
            Culture = CultureInfo.CurrentCulture;
            Encoding = Encoding.UTF8;
            RenderDataAttributes = true;
            Css.ResetToDefault();
        }
        public void Copy(HtmlTableExportSettings copy)
        {
            Minify = copy.Minify;
            HiddenRows = copy.HiddenRows;
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
    public abstract class CssExportSettings
    {
        /// <summary>
        /// If set to true shared css classes used on table elements are included in the css. 
        /// If set to false, these classes has to be included manually. 
        /// Default is true
        /// </summary>
        public bool IncludeSharedClasses { get; set; } = true;
        /// <summary>
        /// If true the normal font will be included in the css. Default is true
        /// </summary>
        public bool IncludeNormalFont { get; set; } = true;

        /// <summary>
        /// Css elements added to the table.
        /// </summary>
        public Dictionary<string, string> AdditionalCssElements
        {
            get;
            internal set;
        }
        /// <summary>
        /// The value used in the stylesheet for an indentation in a cell
        /// </summary>
        public float IndentValue { get; set; } = 2;
        /// <summary>
        /// The unit used in the stylesheet for an indentation in a cell
        /// </summary>
        public string IndentUnit { get; set; } = "em";
        internal void ResetToDefaultInternal()
        {
            AdditionalCssElements = new Dictionary<string, string>()
            {
                { "border-spacing", "0" },
                { "border-collapse", "collapse" },
                { "word-wrap", "break-word"},
                { "white-space", "nowrap"},
            };
            IndentValue = 2;
            IndentUnit = "em";
        }
        internal void CopyInternal(CssExportSettings copy)
        {
            AdditionalCssElements = copy.AdditionalCssElements;
            IndentValue = copy.IndentValue;
            IndentUnit = copy.IndentUnit;
        }
    }
    /// <summary>
    /// Settings for css export for tables
    /// </summary>
    public class CssTableExportSettings : CssExportSettings
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

            Exclude.TableStyle.ResetToDefault();
            Exclude.CellStyle.ResetToDefault();
            base.ResetToDefaultInternal();
        }
        public void Copy(CssTableExportSettings copy)
        {
            IncludeTableStyles = copy.IncludeTableStyles;
            IncludeCellStyles = copy.IncludeTableStyles;

            Exclude.TableStyle.Copy(copy.Exclude.TableStyle);
            Exclude.CellStyle.Copy(copy.Exclude.CellStyle);

            base.CopyInternal(copy);
        }
    }
    /// <summary>
    /// Settings for css export for tables
    /// </summary>
    public class CssRangeExportSettings : CssExportSettings
    {
        public CssRangeExportSettings()
        {
            ResetToDefault();
        }
        public CssExclude CssExclude { get; } = new CssExclude();
        public void ResetToDefault()
        {
            CssExclude.ResetToDefault();
            base.ResetToDefaultInternal();
        }
        public void Copy(CssRangeExportSettings copy)
        {
            CssExclude.Copy(copy.CssExclude);
            base.CopyInternal(copy);
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
