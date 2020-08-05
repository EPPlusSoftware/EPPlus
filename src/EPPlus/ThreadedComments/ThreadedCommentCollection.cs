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

        public ThreadedComment this[int index]
        {
            get
            {
                return _commentList[index];
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

        internal void Add(ThreadedComment comment)
        {
            _commentList.Add(comment);
            if(TopNode.SelectSingleNode("tc:threadedComment[@id='" + comment.Id + "']", NameSpaceManager) == null)
            {
                TopNode.AppendChild(comment.TopNode);
            }
        }
    }
}
