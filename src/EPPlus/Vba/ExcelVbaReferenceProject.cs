/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// A reference to an external VBA project
    /// </summary>
    public class ExcelVbaReferenceProject : ExcelVbaReference
    {
        /// <summary>
        /// Constructor.
        /// Sets ReferenceRecordID to 0x0E
        /// </summary>
        public ExcelVbaReferenceProject()
        {
            ReferenceRecordID = 0x0E;
        }
        /// <summary>
        /// LibIdRelative
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdRelative { get; set; }
        /// <summary>
        /// Major version of the referenced VBA project
        /// </summary>
        public uint MajorVersion { get; set; }
        /// <summary>
        /// Minor version of the referenced VBA project
        /// </summary>
        public ushort MinorVersion { get; set; }
    }
}