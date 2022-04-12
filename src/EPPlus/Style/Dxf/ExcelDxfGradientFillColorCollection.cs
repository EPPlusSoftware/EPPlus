/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A collection of colors and their positions used for a gradiant fill.
    /// </summary>
    public class ExcelDxfGradientFillColorCollection : DxfStyleBase, IEnumerable<ExcelDxfGradientFillColor>
    {
        List<ExcelDxfGradientFillColor> _lst = new List<ExcelDxfGradientFillColor>();
        internal ExcelDxfGradientFillColorCollection(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
            
        }
        /// <summary>
        /// Get the enumerator
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelDxfGradientFillColor> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        /// <summary>
        /// Get the enumerator
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index in the collection</param>
        /// <returns>The color</returns>
        public ExcelDxfGradientFillColor this[int index]
        {
            get
            {
                return (_lst[index]);
            }
        }
        /// <summary>
        /// Gets the first occurance with the color with the specified position
        /// </summary>
        /// <param name="position">The position in percentage</param>
        /// <returns>The color</returns>
        public ExcelDxfGradientFillColor this[double position]
        {
            get
            {
                return (_lst.Find(i => i.Position == position));
            }
        }
        /// <summary>
        /// Adds a RGB color at the specified position
        /// </summary>
        /// <param name="position">The position from 0 to 100%</param>
        /// <returns>The gradient color position object</returns>
        public ExcelDxfGradientFillColor Add(double position)
        {
            if(position < 0 && position > 100)
            {
                throw new ArgumentOutOfRangeException("position","Must be a value between 0 and 100");
            }
            var color = new ExcelDxfGradientFillColor(_styles, position, _callback);
            color.Color.Auto = true;
            _lst.Add(color);
            return color;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _lst.Count;
            }
        }
        internal override string Id
        {
            get
            {
                var id = "";
                foreach (var c in _lst.OrderBy(x=>x.Position))
                {
                    id += c.Id;
                }
                return id;
            }
        }
        /// <summary>
        /// If the style has any value set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return _lst.Count > 0;
            }
        }

        /// <summary>
        /// Remove the style at the index in the collection.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            _lst.RemoveAt(index);
        }
        /// <summary>
        /// Remove the style at the position from the collection.
        /// </summary>
        /// <param name="position"></param>
        public void RemoveAt(double position)
        {
            var item = _lst.Find(i => i.Position == position);
            if(item!=null)
            {
                _lst.Remove(item);
            }
        }
        /// <summary>
        /// Remove the style from the collection
        /// </summary>
        /// <param name="item"></param>
        public void Remove(ExcelDxfGradientFillColor item)
        {
            _lst.Remove(item);
        }
       /// <summary>
       /// Clear all style items from the collection
       /// </summary>
        public override void Clear()
        {
            _lst.Clear();
        }

        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if(_lst.Count>0)
            {
                foreach(var c in _lst)
                {
                    c.CreateNodes(helper, path);
                }
            }
        }
        internal override void SetStyle()
        {
            if (_callback!=null && _lst.Count > 0)
            {
                foreach (var c in _lst)
                {
                    c.SetStyle();
                }
            }
        }
        internal override DxfStyleBase Clone()
        {
            var ret = new ExcelDxfGradientFillColorCollection(_styles, _callback);
            foreach (var c in _lst)
            {
                ret._lst.Add((ExcelDxfGradientFillColor)c.Clone());
            }
            return ret;
        }
    }
}