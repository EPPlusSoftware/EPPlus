using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal class RichTextCollectionBase : IRichTextCollection
    {
        List<IRichText> _list = new List<IRichText>();
        public RichTextCollectionBase()
        {
        }

        public IRichText this[int index] => _list[index];

        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in _list)
                {
                    sb.Append(item.Text);
                }
                return sb.ToString();
            }
        }

        public int Count => _list.Count;

        public IRichText Add(string Text, bool NewParagraph)
        {
            return Insert(_list.Count, Text);
        }

        public IRichText Insert(int index, string Text)
        {
            var rt = new RichTextBase(Text);
            _list.Insert(index, rt);
            return rt;
        }

        public IEnumerator<IRichText> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public void Clear()
        {
            _list.Clear();
        }

        public void RemoveAt(int Index)
        {
            _list.RemoveAt(Index);
        }

        public void Remove(IRichText Item)
        {
            _list.Remove(Item);
        }
    }
}
