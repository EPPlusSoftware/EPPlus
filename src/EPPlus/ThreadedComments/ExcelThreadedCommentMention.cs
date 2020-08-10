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
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// This class represents a mention of a person in a <see cref="ExcelThreadedComment"/>
    /// </summary>
    public class ExcelThreadedCommentMention : XmlHelper
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="nameSpaceManager">Namespace manager of the <see cref="ExcelPackage"/></param>
        /// <param name="topNode">An <see cref="XmlNode"/> representing the mention</param>
        public ExcelThreadedCommentMention(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }

        internal static string NewId()
        {
            var guid = Guid.NewGuid();
            return "{" + guid.ToString().ToUpper() + "}";
        }

        /// <summary>
        /// Index in the <see cref="ExcelThreadedComment"/>s text where the mention starts
        /// </summary>
        public int StartIndex
        {
            get { return GetXmlNodeInt("@startIndex"); }
            set { SetXmlNodeInt("@startIndex", value); }
        }

        /// <summary>
        /// Length of the mention, value for @John Doe would be 9.
        /// </summary>
        public int Length
        {
            get { return GetXmlNodeInt("@length"); }
            set { SetXmlNodeInt("@length", value); }
        }

        /// <summary>
        /// Id of this mention
        /// </summary>
        public string MentionId
        {
            get { return GetXmlNodeString("@mentionId"); }
            set { SetXmlNodeString("@mentionId", value); }
        }

        /// <summary>
        /// Id of the <see cref="ExcelThreadedCommentPerson"/> mentioned
        /// </summary>
        public string MentionPersonId
        {
            get { return GetXmlNodeString("@mentionpersonId"); }
            set { SetXmlNodeString("@mentionpersonId", value); }
        }
    }
}
