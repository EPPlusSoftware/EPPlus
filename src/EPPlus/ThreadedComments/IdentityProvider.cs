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

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// This enum defines the Identity providers for <see cref="ExcelThreadedCommentPerson"/>
    /// as described here: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968
    /// </summary>
    public enum IdentityProvider
    {
        /// <summary>
        /// No provider, Person's userId should be a name
        /// </summary>
        NoProvider,
        /// <summary>
        /// ActiveDirectory, Person's userId should be an ActiveDirectory Security Identifier (SID) as specified here:
        /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/MS-DTYP/cca27429-5689-4a16-b2b4-9325d93e4ba2
        /// </summary>
        ActiveDirectory,
        /// <summary>
        /// Windows Live, Person's userId should be a 64-bit signed decimal that uniquely identifies a user on Windows Live
        /// </summary>
        WindowsLiveId,
        /// <summary>
        /// Office 365. The Person's userId should be a string that uniquely identifies a user. It SHOULD be comprised
        /// of three individual values separated by a &quot;::&quot; delimiter. 
        /// </summary>
        Office365,
        /// <summary>
        /// People Picker, The Persons userId should be an email address provided by People Picker.
        /// </summary>
        PeoplePicker
    }
}
