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
    /// <summary>
    /// Represents a comment in a thread of ThreadedComments
    /// </summary>
    public class ThreadedComment : XmlHelper
    {
        internal ThreadedComment(XmlNode topNode, XmlNamespaceManager namespaceManager, ExcelWorkbook workbook)
            : base(namespaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "text", "mentions" };
            _workbook = workbook;
            Id = NewId();
        }

        private readonly ExcelWorkbook _workbook;

        internal static string NewId()
        {
            var guid = Guid.NewGuid();
            return "{" + guid.ToString().ToUpper() + "}";
        }

        /// <summary>
        /// Indicates whether the Text contains mentions. If so the
        /// Mentions property will contain data about those mentions.
        /// </summary>
        public bool ContainsMentions
        {
            get
            {
                return Mentions != null && Mentions.Any();
            }
        }

        /// <summary>
        /// Address of the cell in the A1 format
        /// </summary>
        public string CellAddress
        {
            get
            {
                return GetXmlNodeString("@ref");
            }
            set
            {
                SetXmlNodeString("@ref", value);
            }
        }

        /// <summary>
        /// Timestamp for when the comment was created
        /// </summary>
        public DateTime DateCreated
        {
            get
            {
                var dt = GetXmlNodeString("@dT");
                if(DateTime.TryParse(dt, out DateTime result))
                {
                    return result;
                }
                throw new InvalidCastException("Could not cast datetime for threaded comment");
            }
            set
            {
                SetXmlNodeString("@dT", value.ToString("yyyy-MM-ddTHH:mm:ss.ff"));
            }
        }

        /// <summary>
        /// Unique id
        /// </summary>
        public string Id
        {
            get
            {
                return GetXmlNodeString("@id");
            }
            set
            {
                SetXmlNodeString("@id", value);
            }
        }

        /// <summary>
        /// Id of the <see cref="ThreadedCommentPerson"/> who wrote the comment
        /// </summary>
        public string PersonId
        {
            get
            {
                return GetXmlNodeString("@personId");
            }
            set
            {
                SetXmlNodeString("@personId", value);
            }
        }

        /// <summary>
        /// Author of the comment
        /// </summary>
        public ThreadedCommentPerson Author
        {
            get
            {
                return _workbook.ThreadedCommentPersons[PersonId];
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string ParentId
        {
            get
            {
                return GetXmlNodeString("@parentId");
            }
            set
            {
                SetXmlNodeString("@parentId", value);
            }
        }

        /// <summary>
        /// Text of the comment
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString("tc:text");
            }
            set
            {
                SetXmlNodeString("tc:text", value);
            }
        }

        /// <summary>
        /// Mentions in this comment. Will return null if no mentions exists.
        /// </summary>
        public ThreadedCommentMentionCollection Mentions
        {
            get
            {
                var mentionsNode = TopNode.SelectSingleNode("tc:mentions", NameSpaceManager);
                if(mentionsNode != null)
                {
                    return new ThreadedCommentMentionCollection(NameSpaceManager, mentionsNode);
                }
                return null;
            }
        }
    }
}
