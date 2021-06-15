/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents a DDE link. This class is read-only.
    /// </summary>
    public class ExcelExternalDdeItem 
    {
        /// <summary>
        /// The name of the DDE link item
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// If the linked object should notify the application when the external data changes.
        /// </summary>
        public bool Advise { get; internal set; }
        /// <summary>
        /// If the linked object is represented by an image.
        /// </summary>
        public bool PreferPicture { get; internal set; }
        /// <summary>
        /// If this is item uses an ole technology.
        /// </summary>
        public bool Ole { get; internal set; }
        /// <summary>
        /// A collection of DDE values
        /// </summary>
        public ExcelExternalDdeValueCollection Values { get; } = new ExcelExternalDdeValueCollection();
    }
}