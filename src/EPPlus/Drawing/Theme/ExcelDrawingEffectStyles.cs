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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Effect;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// The effect styles within the theme
    /// </summary>
    public class ExcelThemeEffectStyles : XmlHelper, IEnumerable<ExcelThemeEffectStyle>
    {
        List<ExcelThemeEffectStyle> _list;
        private readonly ExcelThemeBase _theme;
        internal ExcelThemeEffectStyles(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            _theme = theme;
            _list = new List<ExcelThemeEffectStyle>();
            foreach (XmlNode node in topNode.ChildNodes)
            {
                _list.Add(new ExcelThemeEffectStyle(nameSpaceManager, node, "", null, _theme));
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelThemeEffectStyle> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>The effect style</returns>
        public ExcelThemeEffectStyle this[int index]
        {
            get
            {
                return (_list[index]);
            }
        }
        /// <summary>
        /// Adds a new effect style
        /// </summary>
        /// <returns></returns>
        public ExcelThemeEffectStyle Add()
        {
            var node = TopNode.OwnerDocument.CreateElement("a", "effectStyle", ExcelPackage.schemaMain);
            TopNode.AppendChild(node);
            return new ExcelThemeEffectStyle(NameSpaceManager, TopNode, "", null, _theme);
        }
        /// <summary>
        /// Removes an effect style. The collection must have at least three effect styles.
        /// </summary>
        /// <param name="item">The Item</param>
        public void Remove(ExcelThemeEffectStyle item)
        {
            if (_list.Count == 3)
            {
                throw (new InvalidOperationException("Collection must contain at least 3 items"));
            }

            if (_list.Contains(item))
            {
                _list.Remove(item);
                item.TopNode.ParentNode.RemoveChild(item.TopNode);
            }
        }
        /// <summary>
        /// Remove the effect style at the specified index. The collection must have at least three effect styles.
        /// </summary>
        /// <param name="Index">The index</param>
        public void Remove(int Index)
        {
            if (_list.Count == 3)
            {
                throw (new InvalidOperationException("Collection must contain at least 3 items"));
            }

            if (Index >= _list.Count)
            {
                throw new ArgumentException("Index", "Index out of range");
            }
            _list.Remove(_list[Index]);
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
     }
}
