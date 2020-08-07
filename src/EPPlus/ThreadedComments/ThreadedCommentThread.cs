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
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// Represents a thread of <see cref="ThreadedComment"/>s in a cell on a worksheet. Contains functionality to add and modify these comments.
    /// </summary>
    public class ThreadedCommentThread
    {
        public ThreadedCommentThread(ExcelCellAddress cellAddress, XmlDocument commentsXml, ExcelWorksheet worksheet)
        {
            CellAddress = cellAddress;
            CommentsXml = commentsXml;
            Worksheet = worksheet;
            Comments = new ThreadedCommentCollection(worksheet, commentsXml.SelectSingleNode("tc:ThreadedComments", worksheet.NameSpaceManager));
        }

        /// <summary>
        /// The address of the cell of the comment thread
        /// </summary>
        public ExcelCellAddress CellAddress { get; private set; }

        public ThreadedCommentCollection Comments { get; private set; }

        /// <summary>
        /// The worksheet where this comment thread resides
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get; private set;
        }

        /// <summary>
        /// The raw xml representing this comment thread.
        /// </summary>
        public XmlDocument CommentsXml
        {
            get; private set;
        }

        private void ReplicateThreadToLegacyComment()
        {
            var tc = Comments as IEnumerable<ThreadedComment>;
            if (!tc.Any()) return;
            var tcIndex = 0;
            var commentText = new StringBuilder();
            var authorId = "tc=" + tc.First().Id;
            commentText.AppendLine("This comment reflects a threaded comment in this cell, a feature that might be supported by newer versions of your spreadsheet program (for example later versions of Excel). Any edits will be overwritten if opened in a spreadsheet program that supports threaded comments.");
            commentText.AppendLine();
            foreach(var threadedComment in tc)
            {
                if(tcIndex == 0)
                {
                    commentText.AppendLine("Comment:");
                }
                else
                {
                    commentText.AppendLine("Reply:");
                }
                commentText.AppendLine(threadedComment.Text);
                tcIndex++;
            }
            var comment = Worksheet.Comments[CellAddress];
            if (comment == null)
            {
                Worksheet.Comments.Add(Worksheet.Cells[CellAddress.Address], commentText.ToString(), authorId);
            }
            else
            {
                comment.Text = commentText.ToString();
            }
        }

        /// <summary>
        /// When this method is called the legacy comment representing the thread will be rebuilt.
        /// </summary>
        internal void OnCommentThreadChanged()
        {
            ReplicateThreadToLegacyComment();
        }

        /// <summary>
        /// Adds a <see cref="ThreadedComment"/> to the thread
        /// </summary>
        /// <param name="personId">Id of the author, see <see cref="ThreadedCommentPerson"/></param>
        /// <param name="text">Text of the comment</param>
        public ThreadedComment AddComment(string personId, string text)
        {
            return AddComment(personId, text, true);
        }

        internal ThreadedComment AddComment(string personId, string text, bool replicateLegacyComment)
        {
            Require.That(text).Named("text").IsNotNullOrEmpty();
            Require.That(personId).Named("personId").IsNotNullOrEmpty();
            var parentId = string.Empty;
            if (Comments.Any())
            {
                parentId = Comments.First().Id;
            }
            var xmlNode = CommentsXml.CreateElement("threadedComment", ExcelPackage.schemaThreadedComments);
            CommentsXml.SelectSingleNode("tc:ThreadedComments", Worksheet.NameSpaceManager).AppendChild(xmlNode);
            var newComment = new ThreadedComment(xmlNode, Worksheet.NameSpaceManager, Worksheet.Workbook, this);
            newComment.Id = ThreadedComment.NewId();
            newComment.CellAddress = CellAddress.Address;
            newComment.Text = text;
            newComment.PersonId = personId;
            newComment.DateCreated = DateTime.Now;
            if (!string.IsNullOrEmpty(parentId))
            {
                newComment.ParentId = parentId;
            }
            Comments.Add(newComment);
            if(replicateLegacyComment)
                ReplicateThreadToLegacyComment();
            return newComment;
        }

        internal void AddComment(ThreadedComment comment)
        {
            Comments.Add(comment);
            ReplicateThreadToLegacyComment();
        }

        private void InsertMentions(ThreadedComment comment, params ThreadedCommentPerson[] personsToMention)
        {
            var str = comment.Text;
            var isMentioned = new Dictionary<string, bool>();
            for (var index = 0; index < personsToMention.Length; index++)
            {
                var person = personsToMention[index];
                var format = "{" + index + "}";
                while (str.IndexOf(format) > -1)
                {
                    var placeHolderPos = str.IndexOf("{" + index + "}");
                    var regex = new Regex(@"\{" + index + @"\}");
                    str = regex.Replace(str, "@" + person.DisplayName, 1);

                    // Excel seems to only support one mention per person, so we
                    // add a mention object only for the first occurance per person...
                    if(!isMentioned.ContainsKey(person.Id))
                    {
                        comment.Mentions.AddMention(person, placeHolderPos);
                        isMentioned[person.Id] = true;
                    }
                }
            }
            comment.Mentions.SortAndAddMentionsToXml();
            comment.Text = str;
        }

        /// <summary>
        /// Adds a <see cref="ThreadedComment"/> with mentions in the text to the thread.
        /// </summary>
        /// <param name="personId">Id of the <see cref="ThreadedCommentPerson">autor</see></param>
        /// <param name="textWithFormats">A string with format placeholders - same as in string.Format. Index in these should correspond to an index in the <paramref name="personsToMention"/> array.</param>
        /// <param name="personsToMention">A params array of <see cref="ThreadedCommentPerson"/>. Their DisplayName property will be used to replace the format placeholders.</param>
        /// <returns>The added <see cref="ThreadedComment"/></returns>
        public ThreadedComment AddComment(string personId, string textWithFormats, params ThreadedCommentPerson[] personsToMention)
        {
            var comment = AddComment(personId, textWithFormats, true);
            InsertMentions(comment, personsToMention);
            return comment;
        }

        /// <summary>
        /// Removes a <see cref="ThreadedComment"/> from the thread.
        /// </summary>
        /// <param name="comment">The comment to remove</param>
        /// <returns>true if the comment was removed, otherwise false</returns>
        public bool Remove(ThreadedComment comment)
        {
            if(Comments.Remove(comment))
            {
                ReplicateThreadToLegacyComment();
                return true;
            }
            return false;
        }
    }
}
