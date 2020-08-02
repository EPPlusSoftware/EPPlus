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
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    public class WorksheetThreadedComments
    {
        public WorksheetThreadedComments(ThreadedCommentPersonCollection persons, ExcelWorksheet worksheet)
        {
            Persons = persons;
            _worksheet = worksheet;
            _package = worksheet._package;
            LoadThreads();
        }

        private readonly ExcelWorksheet _worksheet;
        private readonly ExcelPackage _package;
        private readonly List<ThreadedCommentThread> _threads = new List<ThreadedCommentThread>();
        private int _lastId = 0;

        public ThreadedCommentPersonCollection Persons
        {
            get;
            private set;
        }

        public IEnumerable<ThreadedCommentThread> Threads
        {
            get
            {
                return _threads;
            }
        }

        private void LoadThreads()
        {
            var commentRels = _worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaThreadedComment);
            bool isLoaded = false;
            foreach (var commentPart in commentRels)
            {
                var uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
                var part = _package.ZipPackage.GetPart(uri);
                var commentXml = new XmlDocument();
                XmlHelper.LoadXmlSafe(commentXml, part.GetStream());
                var id = commentPart.Id;
                isLoaded = true;
                _threads.Add(new ThreadedCommentThread(commentXml, _worksheet));
            }
            //Create a new document
            if (!isLoaded)
            {
                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
                //Uri = null;
            }
            
        }
    }
}
