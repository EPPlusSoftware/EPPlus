/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Power query permission settings
    /// </summary>
    public class ExcelPowerQueryPermissions 
    {
        /// <summary>
        ///  If the client is allowed to read files created in a newer version of the client. This value is ignored when read, and is written as "false".
        /// </summary>
        internal bool CanEvaluateFuturePackages { get; set; }
        /// <summary>
        /// if the privacy Level settings are used when combining data. See the [MSFT-Support] article <see href="https://support.microsoft.com/en-us/office/set-privacy-levels-power-query-cc3ede4d-359e-4b28-bc72-9bee7900b540">Set Privacy levels (Power Query)</see> for more information.
        /// </summary>
        public bool FirewallEnabled { get; set; }
        /// <summary>
        /// The Privacy Level of the current spreadsheet. See the [MSFT-Support] article <see href="https://support.microsoft.com/en-us/office/set-privacy-levels-power-query-cc3ede4d-359e-4b28-bc72-9bee7900b540">Set Privacy levels (Power Query)</see> for more information.
        /// </summary>
        public eWorkbookGroupType PrivacyLevel { get; set; } = eWorkbookGroupType.None;
    }
    /// <summary>
    /// Permision level for the 
    /// </summary>
    public enum eWorkbookGroupType
    {
        /// <summary>
        /// No privacy settings.
        /// </summary>
        None,
        /// <summary>
        /// Public data source.
        /// </summary>
        Public,
        /// <summary>
        /// Organizational data source.
        /// </summary>
        Organizational,
        /// <summary>
        /// Private data source. The Privacy Level of the current spreadsheet.
        /// </summary>
        SeparatePrivate
    }
}