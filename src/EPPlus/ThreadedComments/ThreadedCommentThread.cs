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
using OfficeOpenXml.FormulaParsing.Utilities;
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
            Comments = new ThreadedCommentCollection(worksheet, commentsXml.SelectSingleNode("tc:ThreadedComments", worksheet.NameSpaceManager));
            if(Comments.Any())
            {
                CellAddress = Comments.First().CellAddress;
            }
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

        /// <summary>
        /// Adds a <see cref="ThreadedComment"/> to the thread
        /// </summary>
        /// <param name="cellAddress">Cell address in A1 format</param>
        /// <param name="personId">Id of the author, see <see cref="ThreadedCommentPerson"/></param>
        /// <param name="text">Text of the comment</param>
        public ThreadedComment AddComment(string cellAddress, string personId, string text)
        {
            Require.That(text).Named("text").IsNotNullOrEmpty();
            Require.That(personId).Named("personId").IsNotNullOrEmpty();
            var parentId = string.Empty;
            if(Comments.Any())
            {
                parentId = Comments.First().Id;
            }
            var xmlNode = CommentsXml.CreateElement("threadedComment", ExcelPackage.schemaThreadedComments);
            CommentsXml.SelectSingleNode("tc:ThreadedComments", Worksheet.NameSpaceManager).AppendChild(xmlNode);
            var newComment = new ThreadedComment(xmlNode, Worksheet.NameSpaceManager, Worksheet.Workbook);
            newComment.CellAddress = cellAddress;
            newComment.Text = text;
            newComment.PersonId = personId;
            newComment.DateCreated = DateTime.Now;
            if(!string.IsNullOrEmpty(parentId))
            {
                newComment.ParentId = parentId;
            }
            Comments.Add(newComment);
            return newComment;
        }

        internal void AddComment(ThreadedComment comment)
        {
            Comments.Add(comment);
        }
    }
}
