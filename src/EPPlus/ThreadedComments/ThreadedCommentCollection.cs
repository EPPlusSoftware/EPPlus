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
    public class ThreadedCommentCollection : IEnumerable<ThreadedComment>
    {
        internal ThreadedCommentCollection(ExcelWorksheet worksheet, XmlDocument commentsXml)
        {
            _package = worksheet._package;
            CommentXml = commentsXml;
            CommentXml.PreserveWhitespace = false;
            NameSpaceManager = worksheet.Workbook.NameSpaceManager;
            Worksheet = worksheet;
            AddCommentsFromXml();
        }

        private readonly ExcelPackage _package;
        public XmlDocument CommentXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
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

        private void AddCommentsFromXml()
        {
            //var lst = new List<IRangeID>();
            foreach (XmlElement node in CommentXml.SelectNodes("tc:ThreadedComments/tc:threadedComment", NameSpaceManager))
            {
                var comment = new ThreadedComment(node, NameSpaceManager, Worksheet.Workbook);
                _commentsIndex[comment.Id] = comment;
                _commentList.Add(comment);
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

        public int Count
        {
            get { return _commentList.Count; }   
        }
    }
}
