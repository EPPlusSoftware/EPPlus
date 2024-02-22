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
    public class ExcelThreadedComment : XmlHelper
    {
        internal ExcelThreadedComment(XmlNode topNode, XmlNamespaceManager namespaceManager, ExcelWorkbook workbook)
            : this(topNode, namespaceManager, workbook, null)
        {
        }

        internal ExcelThreadedComment(XmlNode topNode, XmlNamespaceManager namespaceManager, ExcelWorkbook workbook, ExcelThreadedCommentThread thread)
            : base(namespaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "text", "mentions" };
            _workbook = workbook;
            _thread = thread;
        }

        private readonly ExcelWorkbook _workbook;
        private ExcelThreadedCommentThread _thread;
        internal ExcelThreadedCommentThread Thread
        {
            set
            {
                if (value == null) throw new ArgumentNullException("Thread");
                _thread = value;
            }
        }

        internal static string NewId()
        {
            var guid = Guid.NewGuid();
            return "{" + guid.ToString().ToUpper() + "}";
        }

        /// <summary>
        /// Indicates whether the String contains mentions. If so the
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
        internal string Ref
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
        private ExcelCellAddress _cellAddress=null;
        /// <summary>
        /// The location of the threaded comment
        /// </summary>
        public ExcelCellAddress CellAddress
        {
            get
            {
                if(_cellAddress==null)
                {
                    _cellAddress = new ExcelCellAddress(Ref);
                }
                return _cellAddress;
            }
            internal set
            {
                _cellAddress = value;
                Ref = CellAddress.Address;
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
            internal set
            {
                SetXmlNodeString("@id", value);
            }
        }

        /// <summary>
        /// Id of the <see cref="ExcelThreadedCommentPerson"/> who wrote the comment
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
                _thread.OnCommentThreadChanged();
            }
        }

        /// <summary>
        /// Author of the comment
        /// </summary>
        public ExcelThreadedCommentPerson Author
        {
            get
            {
                return _workbook.ThreadedCommentPersons[PersonId];
            }
        }

        /// <summary>
        /// Id of the first comment in the thread
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
                _thread.OnCommentThreadChanged();
            }
        }

        internal bool? Done
        {
            get
            {
                var val = GetXmlNodeString("@done");
                if(string.IsNullOrEmpty(val))
                {
                    return null;
                }
                if (val == "1") return true;
                return false;
            }
            set
            {
                if(value.HasValue && value.Value)
                {
                    SetXmlNodeInt("@done", 1);
                }
                else if(value.HasValue && !value.Value)
                {
                    SetXmlNodeInt("@done", 0);
                }
                else
                {
                    SetXmlNodeInt("@done", null);
                }
            }
        }

        /// <summary>
        /// String of the comment. To edit the text on an existing comment, use the EditText function.
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString("tc:text");
            }
            internal set
            {
                SetXmlNodeString("tc:text", value);
                _thread.OnCommentThreadChanged();
            }
        }

        /// <summary>
        /// Edit the String of an existing comment
        /// </summary>
        /// <param name="newText"></param>
        public void EditText(string newText)
        {
            Mentions.Clear();
            Text = newText;
            _thread.OnCommentThreadChanged();
        }

        /// <summary>
        /// Edit the String of an existing comment with mentions
        /// </summary>
        /// <param name="newTextWithFormats">A string with format placeholders - same as in string.Format. Index in these should correspond to an index in the <paramref name="personsToMention"/> array.</param>
        /// <param name="personsToMention">A params array of <see cref="ExcelThreadedCommentPerson"/>. Their DisplayName property will be used to replace the format placeholders.</param>
        public void EditText(string newTextWithFormats, params ExcelThreadedCommentPerson[] personsToMention)
        {
            Mentions.Clear();
            MentionsHelper.InsertMentions(this, newTextWithFormats, personsToMention);
            _thread.OnCommentThreadChanged();
        }

        private ExcelThreadedCommentMentionCollection _mentions;

        /// <summary>
        /// Mentions in this comment. Will return null if no mentions exists.
        /// </summary>
        public ExcelThreadedCommentMentionCollection Mentions
        {
            get
            {
                if(_mentions == null)
                {
                    var mentionsNode = TopNode.SelectSingleNode("tc:mentions", NameSpaceManager);
                    if (mentionsNode == null)
                    {
                        mentionsNode = TopNode.OwnerDocument.CreateElement("mentions", ExcelPackage.schemaThreadedComments);
                        TopNode.AppendChild(mentionsNode);
                    }
                    _mentions = new ExcelThreadedCommentMentionCollection(NameSpaceManager, mentionsNode);
                }

                return _mentions;
            }
        }
    }
}
