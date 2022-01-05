namespace OfficeOpenXml
{
    public enum eDataTypeOn
    {
        NoDataTypes,
        OnColumn,
        OnCell
    }
    public abstract class JsonExportSettings
    {
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
        public string ColumnsElementName { get; set; } = "column";
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
    }

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
        /// Write the column array element
        /// </summary>
        public bool WriteColumnsElement { get; set; } = true;
        /// <summary>
        /// Write the table Name attribute
        /// </summary>
        public bool WriteNameAttribute { get; set; } = true;
        /// <summary>
        /// Write the ShowHeader attribute
        /// </summary>
        public bool WriteShowHeaderAttribute { get; set; } = true;
        /// <summary>
        /// Write the ShowHeader attribute
        /// </summary>
        public bool WriteShowTotalsAttribute { get; set; } = true;
    }
}
