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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// A collection of protected ranges in the worksheet.
    ///<seealso cref="ExcelProtection"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    ///<seealso cref="ExcelEncryption"/> 
    /// </summary>
    public class ExcelProtectedRangeCollection : XmlHelper, IEnumerable<ExcelProtectedRange>
    {
        private readonly ExcelWorksheet _ws;
        private List<ExcelProtectedRange> _list = new List<ExcelProtectedRange>();
        private const string _collectionNodePath = "d:protectedRanges";
        private const string _itemNodePath = "protectedRange";
        private XmlElement _collectionNode;
        internal ExcelProtectedRangeCollection(ExcelWorksheet ws)
            : base(ws.NameSpaceManager, ws.TopNode)
        {
            _ws = ws;
            SchemaNodeOrder = ws.SchemaNodeOrder;
            _collectionNode = (XmlElement)GetNode(_collectionNodePath);
            if(_collectionNode!=null)
            {
                foreach (XmlNode node in _collectionNode.ChildNodes)
                {
                    _list.Add(new ExcelProtectedRange(ws.NameSpaceManager, node));
                }
            }
        }

        /// <summary>
        /// Adds a new protected range
        /// </summary>
        /// <param name="name">The name of the protected range</param>
        /// <param name="address">The address within the worksheet</param>
        /// <returns></returns>
        public ExcelProtectedRange Add(string name, ExcelAddress address)
        {
            XmlNode node;
            if (_list.Count==0)
            {
                node = CreateNode($"{_collectionNodePath}/d:{_itemNodePath}");
                _collectionNode = (XmlElement)node.ParentNode;
            }
            else
            {
                node = _collectionNode.OwnerDocument.CreateElement(_itemNodePath, ExcelPackage.schemaMain);
                _collectionNode.AppendChild(node);
            }
            if(_list.Any(x=>x.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
            {
                throw (new InvalidOperationException($"An item with name {name} already exists"));
            }
            var pr = new ExcelProtectedRange(_ws.NameSpaceManager, node) { Name=name, Address=address };
            _list.Add(pr);
            return pr;
        }
        /// <summary>
        /// Clears all protected ranges
        /// </summary>
        public void Clear()
        {
            DeleteNode(_collectionNodePath);
            _list.Clear();
        }
        /// <summary>
        /// Checks if the collection contains a specific item.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Contains(ExcelProtectedRange item)
        {
            return _list.Contains(item);
        }
        /// <summary>
        /// Copies the entire collection to a compatible one-dimensional
        /// array, starting at the specified index of the target array.
        /// </summary>
        /// <param name="array">The array</param>
        /// <param name="arrayIndex">The index</param>
        public void CopyTo(ExcelProtectedRange[] array, int arrayIndex)
        {   
            _list.CopyTo(array, arrayIndex);
        }

        /// <summary>
        /// Numner of items in the collection
        /// </summary>
        public int Count
        {
            get { return _list.Count; }
        }
        /// <summary>
        /// Remove the specified item from the collection
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns></returns>
        public bool Remove(ExcelProtectedRange item)
        {
            item.TopNode.ParentNode.RemoveChild(item.TopNode);            
            var ret = _list.Remove(item);
            if (_list.Count==0)
            {
                _collectionNode.ParentNode.RemoveChild(_collectionNode);
            }
            return ret;
        }

        /// <summary>
        /// Get the index in the collection of the supplied item
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns></returns>
        public int IndexOf(ExcelProtectedRange item)
        {
            return _list.IndexOf(item);
        }

        /// <summary>
        /// Remove the item at the specified indexx
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if(index<0 || index >= _list.Count)
            {
                throw (new IndexOutOfRangeException());
            }
            Remove(_list[index]);
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index to return</param>
        /// <returns></returns>
        public ExcelProtectedRange this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        /// <summary>
        /// Get the enumerator
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator<ExcelProtectedRange> IEnumerable<ExcelProtectedRange>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Get the enumerator
        /// </summary>
        /// <returns>The enumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}
