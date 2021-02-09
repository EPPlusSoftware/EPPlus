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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.ThreadedComments
{
    internal static class MentionsHelper
    {
        /// <summary>
        /// Inserts mentions in the comment text and in the comment
        /// </summary>
        /// <param name="comment"></param>
        /// <param name="textWithFormats">A string with format placeholders with indexes, simlar to string.Format</param>
        /// <param name="personsToMention"><see cref="ExcelThreadedCommentPerson"/>s to mention</param>
        internal static void InsertMentions(ExcelThreadedComment comment, string textWithFormats, params ExcelThreadedCommentPerson[] personsToMention)
        {
            var str = textWithFormats;
            var isMentioned = new Dictionary<string, bool>();
            for (var index = 0; index < personsToMention.Length; index++)
            {
                var person = personsToMention[index];
                var format = "{" + index + "}";
                while (str.IndexOf(format) > -1)
                {
                    var placeHolderPos = str.IndexOf("{" + index + "}", StringComparison.OrdinalIgnoreCase);
                    var regex = new Regex(@"\{" + index + @"\}");
                    str = regex.Replace(str, "@" + person.DisplayName, 1);

                    // Excel seems to only support one mention per person, so we
                    // add a mention object only for the first occurance per person...
                    if (!isMentioned.ContainsKey(person.Id))
                    {
                        comment.Mentions.AddMention(person, placeHolderPos);
                        isMentioned[person.Id] = true;
                    }
                }
            }
            comment.Mentions.SortAndAddMentionsToXml();
            comment.Text = str;
        }
    }
}
