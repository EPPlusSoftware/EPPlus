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
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// A collection of <see cref="ThreadedCommentMention">mentions</see> that occors in a <see cref="ThreadedComment"/>
    /// </summary>
    public sealed class ThreadedCommentMentionCollection : XmlHelper, IEnumerable<ThreadedCommentMention>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="nameSpaceManager">The Namespacemangager of the package</param>
        /// <param name="topNode">The <see cref="XmlNode"/> representing the parent element of the collection</param>
        internal ThreadedCommentMentionCollection(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            LoadMentions();
        }

        private readonly List<ThreadedCommentMention> _mentionList = new List<ThreadedCommentMention>();

        private void LoadMentions()
        {
            foreach(var mentionNode in TopNode.ChildNodes)
            {
                _mentionList.Add(new ThreadedCommentMention(NameSpaceManager, (XmlNode)mentionNode));
            }
        }

        public IEnumerator<ThreadedCommentMention> GetEnumerator()
        {
            return _mentionList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _mentionList.GetEnumerator();
        }

        /// <summary>
        /// Adds a mention
        /// </summary>
        /// <param name="person">The <see cref="ThreadedCommentPerson"/> to mention</param>
        /// <param name="textPosition">Index of the first character of the mention in the text</param>
        internal void AddMention(ThreadedCommentPerson person, int textPosition)
        {
            var elem = TopNode.OwnerDocument.CreateElement("mention", ExcelPackage.schemaThreadedComments);
            var mention = new ThreadedCommentMention(NameSpaceManager, elem);
            mention.MentionId = ThreadedCommentMention.NewId();
            mention.StartIndex = textPosition;
            // + 1 to include the @ prefix...
            mention.Length = person.DisplayName.Length + 1;
            mention.MentionPersonId = person.Id;
            _mentionList.Add(mention);
        }

        /// <summary>
        /// Rebuilds the collection with the elements sorted by the property StartIndex.
        /// </summary>
        internal void SortAndAddMentionsToXml()
        {
            _mentionList.Sort((x, y) => x.StartIndex.CompareTo(y.StartIndex));
            TopNode.RemoveAll();
            _mentionList.ForEach(x => TopNode.AppendChild(x.TopNode));
        }

        /// <summary>
        /// Remove all mentions from the collection
        /// </summary>
        internal void Clear()
        {
            _mentionList.Clear();
            TopNode.RemoveAll();
        }

        public override string ToString()
        {
            return "Count = " + _mentionList.Count;
        }
    }
}
