/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    public partial class ExcelWorksheet
    {
        /// <summary>
        /// Collection containing merged cell addresses
        /// </summary>
        public class MergeCellsCollection : IEnumerable<string>
        {
            internal MergeCellsCollection()
            {

            }
            internal CellStore<int> _cells = new CellStore<int>();
            internal List<string> _list = new List<string>();
            /// <summary>
            /// Indexer for the collection
            /// </summary>
            /// <param name="row">The Top row of the merged cells</param>
            /// <param name="column">The Left column of the merged cells</param>
            /// <returns></returns>
            public string this[int row, int column]
            {
                get
                {
                    int ix = -1;
                    if (_cells.Exists(row, column, ref ix) && ix >= 0 && ix < _list.Count)  //Fixes issue 15075
                    {
                        return _list[ix];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            /// <summary>
            /// Indexer for the collection
            /// </summary>
            /// <param name="index">The index in the collection</param>
            /// <returns></returns>
            public string this[int index]
            {
                get
                {
                    return _list[index];
                }
            }
            internal void Add(ExcelAddressBase address, bool doValidate)
            {
                //Validate
                if (doValidate && Validate(address) == false)
                {
                    throw (new ArgumentException("Can't merge and already merged range"));
                }
                lock (this)
                {
                    var ix = _list.Count;
                    _list.Add(address.Address);
                    SetIndex(address, ix);
                }
            }

            private bool Validate(ExcelAddressBase address)
            {
                int ix = 0;
                if (_cells.Exists(address._fromRow, address._fromCol, ref ix))
                {
                    if (ix >= 0 && ix < _list.Count && _list[ix] != null && address.Address == _list[ix])
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                var cse = new CellStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
                //cells
                while (cse.Next())
                {
                    return false;
                }
                //Entire column
                cse = new CellStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
                while (cse.Next())
                {
                    return false;
                }
                //Entire row
                cse = new CellStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
                while (cse.Next())
                {
                    return false;
                }
                return true;
            }
            internal void SetIndex(ExcelAddressBase address, int ix)
            {
                if (address._fromRow == 1 && address._toRow == ExcelPackage.MaxRows) //Entire row
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        _cells.SetValue(0, col, ix);
                    }
                }
                else if (address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns) //Entire row
                {
                    for (int row = address._fromRow; row <= address._toRow; row++)
                    {
                        _cells.SetValue(row, 0, ix);
                    }
                }
                else
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        for (int row = address._fromRow; row <= address._toRow; row++)
                        {
                            _cells.SetValue(row, col, ix);
                        }
                    }
                }
            }
            /// <summary>
            /// Number of items in the collection
            /// </summary>
            public int Count
            {
                get
                {
                    return _list.Count;
                }
            }
            #region IEnumerable<string> Members

            /// <summary>
            /// Gets the enumerator for the collection
            /// </summary>
            /// <returns>The enumerator</returns>
            public IEnumerator<string> GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion

            #region IEnumerable Members

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion
            internal void Clear(ExcelAddressBase Destination)
            {
                var cse = new CellStoreEnumerator<int>(_cells, Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
                var used = new HashSet<int>();
                while (cse.Next())
                {
                    var v = cse.Value;
                    if (!used.Contains(v) && _list[v] != null)
                    {
                        var adr = new ExcelAddressBase(_list[v]);
                        if (!(Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Inside || Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Equal))
                        {
                            throw (new InvalidOperationException(string.Format("Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {0}", adr._address)));
                        }
                        used.Add(v);
                    }
                }

                _cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
                foreach (var i in used)
                {
                    _list[i] = null;
                }
            }

            internal void CleanupMergedCells()
            {
                _list = _list.Where(x => x != null).ToList();
            }
        }
    }
}
