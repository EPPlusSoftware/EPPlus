using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class RichTextCollectionBase : IRichTextCollection
    {
        List<IRichText> _list = new List<IRichText>();

        public IFontData DefaultFont;
        public IRichTextInfoSimple DefaultRichText;

        /// <summary>
        /// Initalizes using hard-coded defaults
        /// </summary>
        public RichTextCollectionBase() : this(new FontDataDefaults(), new RichTextDefaults())
        {

        }

        /// <summary>
        /// Initializes with user supplied defaults
        /// </summary>
        /// <param name="defaultFont">Default Fallback Font Options</param>
        /// <param name="defaultRichTextOptions">Default RichText Options</param>
        public RichTextCollectionBase(IFontData defaultFont, IRichTextInfoSimple defaultRichTextOptions)
        {
            DefaultFont = defaultFont;
            DefaultRichText = defaultRichTextOptions;
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

        public void Add(IRichText rt)
        {
            Insert(_list.Count, rt);
        }
        public IRichText Insert(int index, IRichText rt)
        {
            _list.Insert(index, rt);
            return rt;
        }

        public IRichText Add(string Text, bool NewParagraph = false)
        {
            return Insert(_list.Count, Text, NewParagraph);
        }

        public IRichText Insert(int index, string Text, bool NewParagraph = false)
        {
            var rt = new RichTextBase(Text, NewParagraph);
            rt.Info = DefaultRichText;
            rt.Info.SetFont(DefaultFont);
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
