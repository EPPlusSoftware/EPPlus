using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// How to set the data type when exporting json.
    /// </summary>
    public enum eDataTypeOn
    {
        /// <summary>
        /// Do not set the data type.
        /// </summary>
        NoDataTypes,
        /// <summary>
        /// Set the data type on the column level.
        /// </summary>
        OnColumn,
        /// <summary>
        /// Set the data type on each cell.
        /// </summary>
        OnCell
    }
    /// <summary>
    /// Base class for settings used when exporting a range or a table as Json.
    /// </summary>
    public abstract class JsonExportSettings
    {
        /// <summary>
        /// If the json is minified when written.
        /// </summary>
        public bool Minify { get; set; } = true;
        /// <summary>
        /// The name of the root element
        /// </summary>
        public abstract string RootElementName { get; set; }
        /// <summary>
        /// Set the dataType attribute depending on the data. The attribute can be set per column or per cell.
        /// </summary>
        public abstract eDataTypeOn AddDataTypesOn { get; set; } 

        /// <summary>
        /// The name of the element containing the columns data
        /// </summary>
        public string ColumnsElementName { get; set; } = "columns";
        /// <summary>
        /// The name of the element containg the rows data
        /// </summary>
        public string RowsElementName { get; set; } = "rows";
        /// <summary>
        /// The name of the element containg the cells data
        /// </summary>
        public string CellsElementName { get; set; } = "cells";
        /// <summary>
        /// Write the uri attribute if an hyperlink exists in a cell
        /// </summary>
        public bool WriteHyperlinks { get; set; } = true;
        /// <summary>
        /// Write the comment attribute if an comment exists in a cell.
        /// </summary>
        public bool WriteComments { get; set; } = true;
        /// <summary>
        /// Encoding for the output
        /// </summary>
        public Encoding Encoding { get; set; } = new UTF8Encoding(false);
        /// <summary>
        /// The CulturInfo used when formatting values.
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;
        /// <summary>
        /// Set if data in worksheet is transposed.
        /// </summary>
        public bool DataIsTransposed { get; set; } = false;
    }
    /// <summary>
    /// Settings used when exporting a range to Json
    /// </summary>
    public class JsonRangeExportSettings : JsonExportSettings
    {
        /// <summary>
        /// The name of the root element
        /// </summary>
        public override string RootElementName { get; set; } = "range";
        /// <summary>
        /// If the first row in the range is the column headers.
        /// The columns array element will be added and the headers will be set using the Name attribute.
        /// </summary>
        public bool FirstRowIsHeader { get; set; } = true;
        /// <summary>
        /// Set the dataType attribute depending on the data. The attribute can be set per column or per cell.
        /// </summary>
        public override eDataTypeOn AddDataTypesOn { get; set; } = eDataTypeOn.OnCell;
    }
    /// <summary>
    /// Settings used when exporting a table to Json
    /// </summary>
    public class JsonTableExportSettings : JsonExportSettings
    {
        /// <summary>
        /// The name of the root element
        /// </summary>
        public override string RootElementName { get; set; } = "table";
        /// <summary>
        /// Set the dataType attribute depending on the data. The attribute can be set per column or per cell.
        /// </summary>
        public override eDataTypeOn AddDataTypesOn { get; set; } = eDataTypeOn.OnColumn;
        /// <summary>
        /// If true the the column array element is written to the output
        /// </summary>
        public bool WriteColumnsElement { get; set; } = true;
        /// <summary>
        /// If true the table Name attribute is written to the output.
        /// </summary>
        public bool WriteNameAttribute { get; set; } = true;
        /// <summary>
        /// If true the ShowHeader attribute is written to the output.
        /// </summary>
        public bool WriteShowHeaderAttribute { get; set; } = true;
        /// <summary>
        /// If true the ShowTotals attribute is written to the output.
        /// </summary>
        public bool WriteShowTotalsAttribute { get; set; } = true;
    }
}
