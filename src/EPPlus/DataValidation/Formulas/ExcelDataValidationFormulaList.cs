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
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaList : ExcelDataValidationFormula, IExcelDataValidationFormulaList
    {
        #region class DataValidationList
        private class DataValidationList : IList<string>, ICollection
        {
            private IList<string> _items = new List<string>();
            private EventHandler<EventArgs> _listChanged;

            public event EventHandler<EventArgs> ListChanged
            {
                add { _listChanged += value; }
                remove { _listChanged -= value; }
            }

            private void OnListChanged()
            {
                if (_listChanged != null)
                {
                    _listChanged(this, EventArgs.Empty);
                }
            }

            #region IList members
            int IList<string>.IndexOf(string item)
            {
                return _items.IndexOf(item);
            }

            void IList<string>.Insert(int index, string item)
            {
                _items.Insert(index, item);
                OnListChanged();
            }

            void IList<string>.RemoveAt(int index)
            {
                _items.RemoveAt(index);
                OnListChanged();
            }

            string IList<string>.this[int index]
            {
                get
                {
                    return _items[index];
                }
                set
                {
                    _items[index] = value;
                    OnListChanged();
                }
            }

            void ICollection<string>.Add(string item)
            {
                _items.Add(item);
                OnListChanged();
            }

            void ICollection<string>.Clear()
            {
                _items.Clear();
                OnListChanged();
            }

            bool ICollection<string>.Contains(string item)
            {
                return _items.Contains(item);
            }

            void ICollection<string>.CopyTo(string[] array, int arrayIndex)
            {
                _items.CopyTo(array, arrayIndex);
            }

            int ICollection<string>.Count
            {
                get { return _items.Count; }
            }

            bool ICollection<string>.IsReadOnly
            {
                get { return false; }
            }

            bool ICollection<string>.Remove(string item)
            {
                var retVal = _items.Remove(item);
                OnListChanged();
                return retVal;
            }

            IEnumerator<string> IEnumerable<string>.GetEnumerator()
            {
                return _items.GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _items.GetEnumerator();
            }
            #endregion

            public void CopyTo(Array array, int index)
            {
                _items.CopyTo((string[])array, index);
            }

            int ICollection.Count
            {
                get { return _items.Count; }
            }

            public bool IsSynchronized
            {
                get { return ((ICollection)_items).IsSynchronized; }
            }

            public object SyncRoot
            {
                get { return ((ICollection)_items).SyncRoot; }
            }
        }
        #endregion

        public ExcelDataValidationFormulaList(string formula, string uid, string sheetName, Action<OnFormulaChangedEventArgs> extListHandler)
            : base(uid, sheetName, extListHandler)
        {
            var values = new DataValidationList();
            values.ListChanged += new EventHandler<EventArgs>(values_ListChanged);
            Values = values;
            _inputFormula = formula;
            SetInitialValues();
        }

        private string _inputFormula;

        private void SetInitialValues()
        {
            var @value = _inputFormula;
            if (!string.IsNullOrEmpty(@value))
            {
                if (@value.StartsWith("\"", StringComparison.OrdinalIgnoreCase) && @value.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
                {
                    @value = @value.Substring(1, @value.Length - 1).TrimEnd('"');
                    var items = @value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < items.Length; i++)
                    {
                        var item = items[i];

                        item = item.Replace("\"\"", "\"");
                        Values.Add(item);
                    }
                }
                else
                {
                    ExcelFormula = @value;
                }
            }
        }

        void values_ListChanged(object sender, EventArgs e)
        {
            if (Values.Count > 0)
            {
                State = FormulaState.Value;
            }
            var valuesAsString = GetValueAsString();
            valuesAsString = valuesAsString?.Trim('\"');
            // Excel supports max 255 characters in this field.
            if (valuesAsString?.Length > 255)
            {
                throw new InvalidOperationException("The total length of a DataValidation list cannot exceed 255 characters");
            }
        }
        public IList<string> Values
        {
            get;
            private set;
        }

        protected override string GetValueAsString()
        {
            if (Values.Count == 0)
            {
                return null;
            }

            var sb = new StringBuilder();

            for (int i = 0; i < Values.Count; i++)
            {
                var val = Values[i];

                if (string.IsNullOrEmpty(val) == false)
                {
                    val = val.Replace("\"", "\"\"");
                }

                if (sb.Length == 0)
                {
                    sb.Append("\"");
                    sb.Append(val);
                }
                else
                {
                    sb.AppendFormat(",{0}", val);
                }
            }
            sb.Append("\"");
            return sb.ToString();
        }

        internal override void ResetValue()
        {
            Values.Clear();
        }
    }
}
