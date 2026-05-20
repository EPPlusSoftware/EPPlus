using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    public interface IRichTextCollection : IEnumerable<IRichText>
    {
        /// <summary>
        /// The full text string of all richtext in the collection
        /// </summary>
        string Text { get; }

        /// <summary>
        /// Collection containing the richtext objects
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        IRichText this[int index] { get; }

        /// <summary>
        /// Number of items in the list
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Add a rich text string
        /// </summary>
        /// <param name="Text">The text to add</param>
        /// <param name="NewParagraph">Adds a new paragraph after the <paramref name="Text"/>. This will add a new line break.</param>
        /// <returns></returns>
        public IRichText Add(string Text, bool NewParagraph);

        /// <summary>
        /// Insert a rich text string at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index at which rich text should be inserted.</param>
        /// <param name="text">The text to insert.</param>
        /// <returns></returns>
        public IRichText Insert(int index, string Text);

        /// <summary>
        /// Clear the collection
        /// </summary>
        public void Clear();

        /// <summary>
        /// Removes an item at the specific index
        /// </summary>
        /// <param name="Index"></param>
        public void RemoveAt(int Index);

        public void Remove(IRichText Item);
    }
}
