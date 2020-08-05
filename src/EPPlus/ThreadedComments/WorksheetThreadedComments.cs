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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// Accessor for <see cref="ThreadedComment"/>s on a <see cref="ExcelWorksheet"/>
    /// </summary>
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
        private readonly Dictionary<string, ThreadedCommentThread> _threads = new Dictionary<string, ThreadedCommentThread>();

        public ThreadedCommentPersonCollection Persons
        {
            get;
            private set;
        }

        /// <summary>
        /// An enumerable of the existing <see cref="ThreadedCommentThread"/>s on the <see cref="ExcelWorksheet">worksheet</see>
        /// </summary>
        public IEnumerable<ThreadedCommentThread> Threads
        {
            get
            {
                return _threads.Values;
            }
        }

        /// <summary>
        /// Number of <see cref="ThreadedCommentThread"/>s on the <see cref="ExcelWorksheet">worksheet</see> 
        /// </summary>
        public int Count
        {
            get { return _threads.Count; }
        }

        /// <summary>
        /// The raw xml for the threaded comments
        /// </summary>
        public XmlDocument CommentsXml
        {
            get; private set;
        }

        private void LoadThreads()
        {
            var commentRels = _worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaThreadedComment);
            foreach (var commentPart in commentRels)
            {
                var uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
                var part = _package.ZipPackage.GetPart(uri);
                CommentsXml = new XmlDocument();
                CommentsXml.PreserveWhitespace = true;
                XmlHelper.LoadXmlSafe(CommentsXml, part.GetStream());
                AddCommentsFromXml();
            }
        }

        private void AddCommentsFromXml()
        {
            foreach (XmlElement node in CommentsXml.SelectNodes("tc:ThreadedComments/tc:threadedComment", _worksheet.NameSpaceManager))
            {
                var comment = new ThreadedComment(node, _worksheet.NameSpaceManager, _worksheet.Workbook);
                var cellAddress = comment.CellAddress.ToUpperInvariant();
                if(!_threads.ContainsKey(cellAddress))
                {
                    var thread = new ThreadedCommentThread(new ExcelCellAddress(comment.CellAddress), CommentsXml, _worksheet);
                    comment.Thread = thread;
                    _threads[cellAddress] = thread;
                }
                _threads[cellAddress].AddComment(comment);
            }
        }

        private void ValidateCellAddress(string cellAddress)
        {
            Require.Argument(cellAddress).IsNotNullOrEmpty("cellAddress");
            if (!ExcelAddress.IsValidCellAddress(cellAddress))
            {
                throw new ArgumentException(cellAddress + " is not a valid cell address. Use A1 format.");
            }
        }

        /// <summary>
        /// Adds a new <see cref="ThreadedCommentThread"/> to the cell.
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <exception cref="ArgumentException">Thrown if there was an existing <see cref="ThreadedCommentThread"/> in the cell.</exception>
        /// <returns>The new, empty <see cref="ThreadedCommentThread"/></returns>
        public ThreadedCommentThread Add(string cellAddress)
        {
            ValidateCellAddress(cellAddress);
            return Add(new ExcelCellAddress(cellAddress));
        }

        public ThreadedCommentThread Add(ExcelCellAddress cellAddress)
        {
            Require.Argument(cellAddress).IsNotNull("cellAddress");
            if (_threads.ContainsKey(cellAddress.Address.ToUpperInvariant()))
            {
                throw new ArgumentException("There is an existing threaded comment thread in cell" + cellAddress);
            }
            if (_worksheet.Comments[cellAddress] != null)
            {
                throw new InvalidOperationException("There is an existing legacy comment/Note in this cell (" + cellAddress + "). See the Worksheet.Comments property. Legacy comments and threaded comments cannot reside in the same cell.");
            }
            CommentsXml = new XmlDocument();
            CommentsXml.PreserveWhitespace = true;
            CommentsXml.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            var thread = new ThreadedCommentThread(cellAddress, CommentsXml, _worksheet);
            _threads[cellAddress.Address.ToUpperInvariant()] = thread;
            return thread;
        }

        /// <summary>
        /// Returns a <see cref="ThreadedCommentThread"/> for the requested <paramref name="cellAddress"/>.
        /// </summary>
        /// <param name="cellAddress">The requested cell address in A1 format</param>
        /// <returns>An existing <see cref="ThreadedCommentThread"/> or null if no thread exists</returns>
        public ThreadedCommentThread this[string cellAddress]
        {
            get
            {
                ValidateCellAddress(cellAddress);
                if (_threads.ContainsKey(cellAddress.ToUpperInvariant())) return _threads[cellAddress.ToUpperInvariant()];
                return null;
            }
        }

        /// <summary>
        /// Returns a <see cref="ThreadedCommentThread"/> for the requested <paramref name="cellAddress"/>.
        /// </summary>
        /// <param name="cellAddress">The requested cell address in A1 format</param>
        /// <returns>An existing <see cref="ThreadedCommentThread"/> or null if no thread exists</returns>
        public ThreadedCommentThread this[ExcelCellAddress cellAddress]
        {
            get
            {
                if (_threads.ContainsKey(cellAddress.Address.ToUpperInvariant())) return _threads[cellAddress.Address.ToUpperInvariant()];
                return null;
            }
        }
    }
}
