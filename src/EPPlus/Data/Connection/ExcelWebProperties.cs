/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography.Xml;
using System.Security.Policy;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Connection properties for web queries.
    /// </summary>
    public class ExcelWebProperties
    {
        /// <summary>
        /// True if the web query source is XML (versus HTML), otherwise false.
        /// </summary>
        public bool IsXml { get; set; }
        /// <summary>
        /// If XML source data should be imported instead of the HTML table itself. 
        /// Used when a web query exists to an HTML table with the following attribute.
        /// &lt;TABLE … o:WebQuerySourceHRef=”http://…” … &gt; … &lt;/TABLE&gt;
        /// </summary>
        public bool IsXmlSourceData { get; set; }
        /// <summary>
        /// If data contained within HTML &lt;PRE&gt; tags in the web page is parsed into columns when you import the page into a query table.
        /// </summary>
        public bool ParsePRE { get; set; }
        /// <summary>
        /// If consecutive delimiters should be treated as just one delimiter
        /// </summary>
        public bool Consecutive { get; set; }
        /// <summary>
        /// Whether to parse all tables inside a PRE block with the same width settings as the first row.
        /// </summary>
        public bool FirstRow { get; set; }
        /// <summary>
        /// If the web query is created in Excel 97
        /// </summary>
        public bool IsExcel97 { get; set; }
        /// <summary>
        /// If the web query is created or refreshed in Excel 2000 or later.
        /// </summary>
        public bool IsExcel2000 { get; set; }
        /// <summary>
        /// If dates should be imported into cells in the worksheet as text rather than dates.
        /// </summary>
        public bool TextDates { get; set; }
        /// <summary>
        /// The URL to use to refresh external data
        /// </summary>
        public string Url { get; set; }
        /// <summary>
        /// Returns or sets the string used with the post method of inputting data into a web server to return data from a web query.
        /// </summary>
        public string Post { get; set; }
        /// <summary>
        /// The URL of the user-facing web page showing the web query data. 
        /// This URL is persisted in the case that <see cref="IsXmlSourceData"/> = "true" and url has been redirected to reference an XML file.
        /// Then the user-facing page can be shown in the UI, and the XML data can be retrieved behind the scenes.
        /// </summary>
        public string EditPage { get; set; }
        /// <summary>
        /// If web queries should only work on HTML tables.
        /// </summary>
        public bool HtmlTables { get; set; }
        /// <summary>
        /// How to handle formatting from the HTML source when bringing web query data into the worksheet.Relevant when <see cref="IsXmlSourceData"/>is True
        /// </summary>
        public eHtmlFormatingHandling HtmlFormat { get; set; }
        /// <summary>
        /// A list of all tables in the HTML document that can be imported using the web query.
        /// </summary>
        public List<ExcelHtmlTableReference> Tables { get; } = new List<ExcelHtmlTableReference>();
    }
}