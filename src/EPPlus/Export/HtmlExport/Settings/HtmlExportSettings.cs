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
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Base class for HTML export for ranges and tables.
    /// </summary>
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
        /// Use this property to set the name of the html data-* attribute that contains the raw value.
        /// Default value is "value" which means that the name of the attribute is "data-value".
        /// To change the name to "data-x", set this property to "x"
        /// </summary>
        public string DataValueAttributeName
        {
            get; set;
        } = "value";
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
        /// <item><term>Cell styles</term><term>{StyleClassPrefix}{CellStyleClassName}{index}</term></item>
        /// <item><term>Hidden Row{StyleClassPrefix}hidden</term></item>
        /// <item><term>Alignment Left </term><term>{StyleClassPrefix}al</term></item>
        /// <item><term>Alignment Right</term><term>{StyleClassPrefix}ar</term></item>
        /// <item><term>Default column width</term><term>{StyleClassPrefix}dcw</term></item>
        /// <item><term>Default row height</term><term>{StyleClassPrefix}drh</term></item>
        /// <item><term>Image content and position</term><term>{StyleClassPrefix}image-{imageName}</term></item>
        /// <item><term>Properties for an image, for example position or border settings</term><term>{StyleClassPrefix}image-prop-{imageName}</term></item>
        /// <item><term>Alignment for cells containing an image</term><term>{StyleClassPrefix}image-cell}</term></item>
        /// </list>
        /// <seealso cref="CellStyleClassName"/>
        /// <seealso cref="Pictures"/>
        /// </summary>
        public string StyleClassPrefix { get; set; } = "epp-";
        /// <summary>
        /// The name of the classes used for cell styles. 
        /// The name will be prefixed with the <see cref="StyleClassPrefix"/> and suffixed with the cell style index.
        /// </summary>
        public string CellStyleClassName { get; set; } = "s";
		/// <summary>
		/// The name of the classes used for cell styles. 
		/// The name will be prefixed with the <see cref="StyleClassPrefix"/> and suffixed with the cell style index.
		/// </summary>
		public string DxfStyleClassName { get; set; } = "dxf";

        /// <summary>
        /// The name of the classes used for cell styles. 
        /// The name will be prefixed with the <see cref="DxfStyleClassName"/> and suffixed with the conditional formatting style index.
        /// </summary>
        public string ConditionalFormattingClassName { get; set; } = "cf";

        /// <summary>
        /// The name of the classes used for cell styles. 
        /// The name will be prefixed with the <see cref="StyleClassPrefix"/>.
        /// </summary>
        public string DatabarPrefix { get; set; } = "db";

        /// <summary>
        /// The name of the classes used to store icon images. 
        /// The name will be prefixed with the <see cref="StyleClassPrefix"/> and suffixed with the icon name.
        /// </summary>
        public string IconPrefix { get; set; } = "ic";

        /// <summary>
        /// If picture drawings will be included. Default is true.
        /// </summary>
        public HtmlPictureSettings Pictures
        {
            get;
        } = new HtmlPictureSettings();

        /// <summary>
        /// If set to true classes that identifies Excel table styling will be included in the html. Default value is true.
        /// </summary>
        public bool IncludeCssClassNames { get; set; } = true;

        /// <summary>
        /// Set the target attribute for hyperlinks (a elements) in the exported html.
        /// Can be null/empty (no target attribute), _blank, _top, _self, _parent or a frame-name.
        /// </summary>
        public string HyperlinkTarget { get; set; }

        /// <summary>
        /// If and how table styles should export, if the range is a table.
        /// The range must be the same as the table range.
        /// </summary>
        public eHtmlRangeTableInclude TableStyle
        {
            get;
            set;
        } = eHtmlRangeTableInclude.Include;
        /// <summary>
        /// If true data-* attributes will be rendered
        /// </summary>
        public bool RenderDataAttributes { get; set; } = true;
        /// <summary>
        /// If true, data types are renedered on the header objects.
        /// </summary>
        public bool RenderDataTypes { get; set; } = true;
        /// <summary>
        /// If true conditionalFormattings will be rendered
        /// </summary>
        public bool RenderConditionalFormattings { get; set; } = true;
    }
}
