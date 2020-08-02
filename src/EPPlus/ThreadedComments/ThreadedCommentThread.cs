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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    public class ThreadedCommentThread
    {
        public ThreadedCommentThread(XmlDocument commentsXml, ExcelWorksheet worksheet)
        {
            CommentsXml = commentsXml;
            Worksheet = worksheet;
            Comments = new ThreadedCommentCollection(worksheet, commentsXml);
            CellAddress = Comments.First().CellAddress;
        }
        public string CellAddress { get; private set; }

        public ThreadedCommentCollection Comments { get; private set; }

        public ExcelWorksheet Worksheet
        {
            get; private set;
        }

        public XmlDocument CommentsXml
        {
            get; private set;
        }
    }
}
