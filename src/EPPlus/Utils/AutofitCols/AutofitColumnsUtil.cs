using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Utils.AutofitCols
{
    internal class AutofitColumnsUtil
    {
        public AutofitColumnsUtil(ExcelRangeBase range)
        {
            _worksheet = range._worksheet;
            _range = range;
            _fromCol = range._fromCol;
            _toCol = range._toCol;
            _fromRow = range._fromRow;
            _toRow = range._toRow;
        }

        private readonly ExcelWorksheet _worksheet;
        private readonly ExcelRangeBase _range;
        private readonly int _fromCol, _toCol;
        private readonly int _fromRow, _toRow;

        internal void AutofitColumns(double MinimumWidth, double MaximumWidth)
        {
            if (_worksheet.Dimension == null)
            {
                return;
            }
            if (_fromCol < 1 || _fromRow < 1)
            {
                SetToSelectedRange();
            }
            var fontCache = new FontCache();

            bool doAdjust = _worksheet._package.DoAdjustDrawings;
            _worksheet._package.DoAdjustDrawings = false;
            var drawWidths = _worksheet.Drawings.GetDrawingWidths();

            var fromCol = _fromCol > _worksheet.Dimension._fromCol ? _fromCol : _worksheet.Dimension._fromCol;
            var toCol = _toCol < _worksheet.Dimension._toCol ? _toCol : _worksheet.Dimension._toCol;

            if (fromCol > toCol) return; //Issue 15383

            if (_range.Addresses == null)
            {
                SetMinWidth(MinimumWidth, fromCol, toCol);
            }
            else
            {
                foreach (var addr in _range.Addresses)
                {
                    fromCol = addr._fromCol > _worksheet.Dimension._fromCol ? addr._fromCol : _worksheet.Dimension._fromCol;
                    toCol = addr._toCol < _worksheet.Dimension._toCol ? addr._toCol : _worksheet.Dimension._toCol;
                    SetMinWidth(MinimumWidth, fromCol, toCol);
                }
            }

            //Get any autofilter to widen these columns
            var afAddr = new List<ExcelAddressBase>();
            if (_worksheet.AutoFilterAddress != null)
            {
                afAddr.Add(new ExcelAddressBase(_worksheet.AutoFilterAddress._fromRow,
                                                    _worksheet.AutoFilterAddress._fromCol,
                                                    _worksheet.AutoFilterAddress._fromRow,
                                                    _worksheet.AutoFilterAddress._toCol));
                afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
            }
            foreach (var tbl in _worksheet.Tables)
            {
                if (tbl.AutoFilterAddress != null)
                {
                    afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._fromCol,
                                                                            tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._toCol));
                    afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
                }
            }

            // if more than 2000 cells, split to 4 tasks
            var tasks = new List<MeasurementTask>();
            var nTasks = _range.Count() > 2000 ? 4 : 1;
            var nCellsPerTask = _range.Count() / nTasks;

            for(var x = 0; x < nTasks; x++)
            {
                var range = _range.Skip(x * nCellsPerTask).Take(nCellsPerTask);
                var measurementTask = new MeasurementTask(_range, fontCache, afAddr);
                tasks.Add(measurementTask);
            }

            // switch between these two if statements to enable/disable multithreading
            //if(true)
            if(nTasks == 1)
            {
                // execute tasks in sequence (on a single thread)
                tasks.ForEach(t => t.Execute());
            }
            else
            {
                // create threads
                var threads = new List<Thread>();
                tasks.ForEach(t =>
                {
                    var thread = new Thread(new ThreadStart(t.Execute));
                    threads.Add(thread);
                    thread.Start();
                });
                Thread.Sleep(10);
                // wait for threads to finish...
                threads.ForEach(t => t.Join());
            }

            // collect data from the tasks
            var result = new Dictionary<int, double>();
            foreach(var task in tasks)
            {
                var taskResults = task.GetResult();
                foreach(var col in taskResults.Keys)
                {
                    if(!result.ContainsKey(col))
                    {
                        result[col] = taskResults[col];
                    }
                    else if(taskResults[col] > result[col])
                    {
                        result[col] = taskResults[col];
                    }
                }
            }

            // set width on columns
            foreach (var col in result.Keys)
            {
                var colWidth = result[col];
                _worksheet.Column(col).Width = colWidth > MaximumWidth ? MaximumWidth : colWidth;
            }
            _worksheet.Drawings.AdjustWidth(drawWidths);
            _worksheet._package.DoAdjustDrawings = doAdjust;
        }

        private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
        {
            var iterator = new CellStoreEnumerator<ExcelValue>(_worksheet._values, 0, fromCol, 0, toCol);
            var prevCol = fromCol;
            foreach (ExcelValue val in iterator)
            {
                var col = (ExcelColumn)val._value;
                if (col.Hidden) continue;
                col.Width = minimumWidth;
                if (_worksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
                {
                    var newCol = _worksheet.Column(prevCol);
                    newCol.ColumnMax = col.ColumnMin - 1;
                    newCol.Width = minimumWidth;
                }
                prevCol = col.ColumnMax + 1;
            }
            if (_worksheet.DefaultColWidth > minimumWidth && prevCol < toCol)
            {
                var newCol = _worksheet.Column(prevCol);
                newCol.ColumnMax = toCol;
                newCol.Width = minimumWidth;
            }
        }

        private void SetToSelectedRange()
        {
            if (_worksheet.View.SelectedRange == "")
            {
                _range.Address = "A1";
            }
            else
            {
                _range.Address = _worksheet.View.SelectedRange;
            }
        }
    }
}
