/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/29/2020         EPPlus Software AB       Threaded comments
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// This class represents an enumerable of <see cref="ThreadedComment"/>s.
    /// </summary>
    public class ThreadedCommentCollection : XmlHelper, IEnumerable<ThreadedComment>
    {
        internal ThreadedCommentCollection(ExcelWorksheet worksheet, XmlNode topNode)
            : base(worksheet.NameSpaceManager, topNode)
        {
            _package = worksheet._package;
            Worksheet = worksheet;
        }

        private readonly ExcelPackage _package;
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }

        private readonly Dictionary<string, ThreadedComment> _commentsIndex = new Dictionary<string, ThreadedComment>();
        private readonly List<ThreadedComment> _commentList = new List<ThreadedComment>();

        /// <summary>
        /// A reference to the worksheet object
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get;
            set;
        }

        /// <summary>
        /// Returns a <see cref="ThreadedComment"/> by its index
        /// </summary>
        /// <param name="index">Index in this collection</param>
        /// <returns>The <see cref="ThreadedComment"/> at the requested <paramref name="index"/></returns>
        /// <exception cref="ArgumentOutOfRangeException">If the <paramref name="index"/> falls out of range</exception>
        public ThreadedComment this[int index]
        {
            get
            {
                return _commentList[index];
            }
        }

        /// <summary>
        /// Returns a <see cref="ThreadedComment"/> by its <paramref name="id"/>
        /// </summary>
        /// <param name="id">Id of the requested <see cref="ThreadedComment"/></param>
        /// <returns>The requested <see cref="ThreadedComment"/></returns>
        /// <exception cref="ArgumentException">If the requested <paramref name="id"/> was not present.</exception>
        public ThreadedComment this[string id]
        {
            get
            {
                if(!_commentsIndex.ContainsKey(id))
                {
                    throw new ArgumentException("Id " + id + " was not present in the comments.");
                }
                return _commentsIndex[id];
            }
        }



        public IEnumerator<ThreadedComment> GetEnumerator()
        {
            return _commentList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _commentList.GetEnumerator();
        }

        /// <summary>
        /// Number of <see cref="ThreadedComment"/>s
        /// </summary>
        public int Count
        {
            get { return _commentList.Count; }   
        }

        private void RebuildIndex()
        {
            _commentsIndex.Clear();
            foreach(var comment in _commentList)
            {
                _commentsIndex[comment.Id] = comment;
            }
        }

        internal void Add(ThreadedComment comment)
        {
            _commentList.Add(comment);
            if(TopNode.SelectSingleNode("tc:threadedComment[@id='" + comment.Id + "']", NameSpaceManager) == null)
            {
                TopNode.AppendChild(comment.TopNode);
            }
            RebuildIndex();
        }

        internal bool Remove(ThreadedComment comment)
        {
            _commentList.Remove(comment);
            if (TopNode.SelectSingleNode("tc:threadedComment[@id='" + comment.Id + "']", NameSpaceManager) != null)
            {
                TopNode.RemoveChild(comment.TopNode);
                return true;
            }
            RebuildIndex();
            return false;
        }
    }
}
