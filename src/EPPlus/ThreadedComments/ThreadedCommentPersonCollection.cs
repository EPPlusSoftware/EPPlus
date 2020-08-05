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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    public class ThreadedCommentPersonCollection : IEnumerable<ThreadedCommentPerson>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="workbook">The <paramref name="workbook"/> where the <see cref="ThreadedCommentPerson"/> occurs</param>
        public ThreadedCommentPersonCollection(ExcelWorkbook workbook)
        {
            _workbook = workbook;
            if(workbook._package.ZipPackage.PartExists(workbook.PersonsUri))
            {
                PersonsXml = workbook._package.GetXmlFromUri(workbook.PersonsUri);
                // lägg upp personerna i listan, loopa på noderna
                var listNode = PersonsXml.DocumentElement;
                foreach(var personNode in listNode.ChildNodes)
                {
                    var person = new ThreadedCommentPerson(workbook.NameSpaceManager, (XmlNode)personNode);
                    _personList.Add(person);
                }
            }
            else
            {
                PersonsXml = new XmlDocument();
                PersonsXml.LoadXml("<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            }
        }

        private readonly ExcelWorkbook _workbook;
        private readonly List<ThreadedCommentPerson> _personList = new List<ThreadedCommentPerson>();

        public XmlDocument PersonsXml { get; private set; }

        /// <summary>
        /// Number of <see cref="ThreadedCommentPerson"/>s in the collection
        /// </summary>
        public int Count 
        { 
            get 
            {
                return _personList.Count;
            } 
        }

        /// <summary>
        /// Returns the <see cref="ThreadedCommentPerson"/> by its index
        /// </summary>
        /// <param name="index">The requested index</param>
        /// <returns>The <see cref="ThreadedCommentPerson"/> at the requested index</returns>
        public ThreadedCommentPerson this[int index]
        {
            get
            {
                return _personList[index];
            }
        }

        /// <summary>
        /// Returns a <see cref="ThreadedCommentPerson"/> by its id
        /// </summary>
        /// <param name="id">The Id of the Person</param>
        /// <returns>A <see cref="ThreadedCommentPerson"/> with the requested <paramref name="id"/> or null</returns>
        public ThreadedCommentPerson this[string id]
        {
            get
            {
                return _personList.FirstOrDefault(x => x.Id == id);
            }
        }

        /// <summary>
        /// Finds a <see cref="ThreadedCommentPerson"/> that <paramref name="match"/> a certain criteria
        /// </summary>
        /// <param name="match">The criterias</param>
        /// <returns>A matching <see cref="ThreadedCommentPerson"/></returns>
        public ThreadedCommentPerson Find(Predicate<ThreadedCommentPerson> match)
        {
            return _personList.Find(match);
        }

        /// <summary>
        /// Finds a number of <see cref="ThreadedCommentPerson"/>'s that matches a certain criteria.
        /// </summary>
        /// <param name="match">The criterias</param>
        /// <returns>An enumerable of matching <see cref="ThreadedCommentPerson"/>'s</returns>
        public IEnumerable<ThreadedCommentPerson> FindAll(Predicate<ThreadedCommentPerson> match)
        {
            return _personList.FindAll(match);
        }

        /// <summary>
        /// Creates and adds a new <see cref="ThreadedCommentPerson"/> to the workbooks list of persons. A unique Id for the person will be generated and set.
        /// The userId will be the same as the <paramref name="displayName"/> and identityProvider will be set to <see cref="IdentityProvider.NoProvider"/>
        /// </summary>
        /// <param name="displayName">The display name of the added <see cref="ThreadedCommentPerson"/></param>
        public ThreadedCommentPerson Add(string displayName)
        {
            return Add(displayName, displayName, IdentityProvider.NoProvider);
        }

        /// <summary>
        /// Creates and adds a new <see cref="ThreadedCommentPerson"/> to the workbooks list of persons. A unique Id for the person will be generated and set.
        /// </summary>
        /// <param name="displayName">The display name of the added <see cref="ThreadedCommentPerson"/></param>
        /// <param name="userId">A string representing the userId of the <paramref name="identityProvider"/></param>
        /// <param name="identityProvider">The <see cref="IdentityProvider"/> from which the <see cref="ThreadedCommentPerson"/> originates</param>
        /// <returns>The added <see cref="ThreadedCommentPerson"/></returns>
        public ThreadedCommentPerson Add(string displayName, string userId, IdentityProvider identityProvider)
        {
            return Add(displayName, userId, identityProvider, ThreadedCommentPerson.NewId());
        }

        /// <summary>
        /// Creates and adds a new <see cref="ThreadedCommentPerson"/> to the workbooks list of persons
        /// </summary>
        /// <param name="displayName">The display name of the added <see cref="ThreadedCommentPerson"/></param>
        /// <param name="userId">A string representing the userId of the <paramref name="identityProvider"/></param>
        /// <param name="identityProvider">The <see cref="IdentityProvider"/> from which the <see cref="ThreadedCommentPerson"/> originates</param>
        /// <param name="id">Id of the <see cref="ThreadedCommentPerson"/></param>
        /// <returns>The added <see cref="ThreadedCommentPerson"/></returns>
        public ThreadedCommentPerson Add(string displayName, string userId, IdentityProvider identityProvider, string id)
        {
            var personsNode = PersonsXml.CreateElement("person", ExcelPackage.schemaThreadedComments);
            PersonsXml.DocumentElement.AppendChild(personsNode);
            var p = new ThreadedCommentPerson(_workbook.NameSpaceManager, personsNode);
            p.DisplayName = displayName;
            p.Id = id;
            p.UserId = userId;
            p.ProviderId = identityProvider;
            _personList.Add(p);
            return p;
        }

        public IEnumerator<ThreadedCommentPerson> GetEnumerator()
        {
            return _personList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _personList.GetEnumerator();
        }

        internal void Add(ThreadedCommentPerson person)
        {
            _personList.Add(person);
        }

        /// <summary>
        /// Removes a <see cref="ThreadedCommentPerson"/> from the collection
        /// </summary>
        /// <param name="person"></param>
        public void Remove(ThreadedCommentPerson person)
        {
            var node = PersonsXml.DocumentElement.SelectSingleNode("/person[id='" + person.Id + "']");
            if(node != null)
            {
                PersonsXml.DocumentElement.RemoveChild(node);
            }
            _personList.Remove(person);
        }

        /// <summary>
        /// Removes all persons from the collection
        /// </summary>
        public void Clear()
        {
            PersonsXml.DocumentElement.RemoveAll();
            _personList.Clear();
        }
    }
}
