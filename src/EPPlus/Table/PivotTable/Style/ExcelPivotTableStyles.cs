using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableStyle
    {
        ExcelStyles _styles;
        internal ExcelPivotTableStyle(ExcelPivotTable pt)
        {
            _styles = pt.WorkSheet.Workbook.Styles;
            Areas = new ExcelPivotTableAreaStyleCollection(pt);
        }
        ExcelPivotTableAreaStyle _all = null;
        public ExcelPivotTableAreaStyle All
        {
            get
            {
                if(_all==null)
                {
                    _all = Areas.Add(ePivotAreaType.All);
                    _all.DataOnly = false;
                    _all.LabelOnly = false;
                    _all.GrandRow = true;
                    _all.GrandColumn = true;
                }
                return _all;
            }
        }
        ExcelPivotTableAreaStyle _labels = null;
        public ExcelPivotTableAreaStyle Labels
        {
            get
            {
                if (_labels == null)
                {
                    _labels = Areas.Add(ePivotAreaType.All);
                    _labels.DataOnly = false;
                    _labels.LabelOnly = true;
                    _labels.GrandRow = true;
                    _labels.GrandColumn = true;
                }
                return _labels;
            }
        }
        ExcelPivotTableAreaStyle _columnLabels = null;
        public ExcelPivotTableAreaStyle ColumnLabels
        {
            get
            {
                if (_columnLabels == null)
                {
                    _columnLabels = Areas.Add();
                    _columnLabels.DataOnly = false;
                    _columnLabels.LabelOnly = true;
                    _columnLabels.GrandRow = false;
                    _columnLabels.GrandColumn = true;
                }
                return _columnLabels;
            }
        }
        ExcelPivotTableAreaStyle _data = null;
        public ExcelPivotTableAreaStyle Data
        {
            get
            {
                if (_data == null)
                {
                    _data = Areas.Add();
                    _data.DataOnly = true;
                    _data.LabelOnly = false;
                    _data.GrandRow = false;
                    //_data.FieldPosition = 0;
                    //_data.FieldIndex = 0;
                    _data.GrandColumn = false;
                }
                return _data;
            }
        }
        ExcelPivotTableAreaStyle _grandRowData = null;
        public ExcelPivotTableAreaStyle GrandRowData
        {
            get
            {
                if (_grandRowData == null)
                {
                    _grandRowData = Areas.Add();
                    _grandRowData.DataOnly = true;
                    _grandRowData.LabelOnly = false;
                    _grandRowData.GrandRow = true;
                    _grandRowData.GrandColumn = false;
                    _grandRowData.FieldPosition = 0;
                    _grandRowData.FieldIndex = 0;
                    _grandRowData.CollapsedLevelsAreSubtotals = true;
                }
                return _grandRowData;
            }
        }
        ExcelPivotTableAreaStyle _grandColumnData = null;
        public ExcelPivotTableAreaStyle GrandColumnData
        {
            get
            {
                if (_grandColumnData == null)
                {
                    _grandColumnData = Areas.Add();
                    _grandColumnData.DataOnly = true;
                    _grandColumnData.LabelOnly = false;
                    _grandColumnData.GrandRow = false;
                    _grandColumnData.GrandColumn = true;
                    _grandColumnData.FieldPosition = 0;
                    _grandColumnData.FieldIndex = 0;
                    _grandColumnData.CollapsedLevelsAreSubtotals = true;
                }
                return _grandColumnData;
            }
        }

        ExcelPivotTableAreaStyle _origin = null;
        public ExcelPivotTableAreaStyle Origin
        {
            get
            {
                if (_origin == null)
                {
                    _origin = Areas.Add(ePivotAreaType.Origin);
                    _origin.DataOnly = false;
                    _origin.LabelOnly = true;
                    _origin.FieldIndex = 0;
                    _origin.FieldPosition = 0;

                }
                return _origin;
            }
        }
        ExcelPivotTableAreaStyle _columnHeaders = null;
        public ExcelPivotTableAreaStyle ColumnHeaders
        {
            get
            {
                if (_columnHeaders == null)
                {
                    _columnHeaders = Areas.Add(ePivotAreaType.FieldButton, ePivotTableAxis.RowAxis);
                    _columnHeaders.DataOnly = false;
                    _columnHeaders.LabelOnly = true;
                    _columnHeaders.FieldIndex = 0;
                    _columnHeaders.FieldPosition = 0;
                }
                return _columnHeaders;
            }
        }
        ExcelPivotTableAreaStyle _grandColumnHeaders = null;
        public ExcelPivotTableAreaStyle GrandColumnHeaders
        {
            get
            {
                if (_grandColumnHeaders == null)
                {
                    _grandColumnHeaders = Areas.Add(ePivotAreaType.FieldButton);
                    _grandColumnHeaders.DataOnly = false;
                    _grandColumnHeaders.LabelOnly = true;
                    _grandColumnHeaders.FieldIndex = 0;
                    _grandColumnHeaders.FieldPosition = 0;
                    //_grandColumnHeaders.GrandRow = true;
                }

                return _grandColumnHeaders;
            }
        }


        public ExcelPivotTableAreaStyleCollection Areas
        {
            get;
        } 
    }
}
