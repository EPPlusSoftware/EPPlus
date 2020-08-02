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
    public class ThreadedCommentMentionCollection : XmlHelper, IEnumerable<ThreadedCommentMention>
    {
        public ThreadedCommentMentionCollection(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
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
    }
}
