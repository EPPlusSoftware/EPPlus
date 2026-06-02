using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    public class RichTextCollectionBase : IRichTextCollection
    {
        List<IRichTextFormatBase> _list = new List<IRichTextFormatBase>();

        public IRichTextFormatBase DefaultRichText = new RichTextFormatSimple();

        /// <summary>
        /// Initalizes using hard-coded defaults
        /// </summary>
        public RichTextCollectionBase() : this(new RichTextFormatSimple())
        {
        }

        /// <summary>
        /// Initializes with user supplied defaults
        /// </summary>
        /// <param name="defaultFont">Default Fallback Font Options</param>
        /// <param name="defaultRichTextOptions">Default RichText Options</param>
        public RichTextCollectionBase(IRichTextFormatBase defaultRichTextOptions)
        {
            DefaultRichText = defaultRichTextOptions;
        }

        public IRichTextFormatBase this[int index] => _list[index];

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

        public void Add(IRichTextFormatBase rt)
        {
            Insert(_list.Count, rt);
        }
        public IRichTextFormatBase Insert(int index, IRichTextFormatBase rt)
        {
            _list.Insert(index, rt);
            return rt;
        }

        public IRichTextFormatBase Add(string Text, bool NewParagraph = false)
        {
            return Insert(_list.Count, Text, NewParagraph);
        }

        public IRichTextFormatBase Insert(int index, string Text, bool NewParagraph = false)
        {
            var rt = new RichTextFormatBase();
            rt.Text = Text;
            rt.SetFont(DefaultRichText);
            _list.Insert(index, rt);
            return rt;
        }

        public IEnumerator<IRichTextFormatBase> GetEnumerator()
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

        public void Remove(IRichTextFormatBase Item)
        {
            _list.Remove(Item);
        }
    }
}
