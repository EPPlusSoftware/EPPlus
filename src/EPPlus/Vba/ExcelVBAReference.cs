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
    /// A VBA reference
    /// </summary>
    public class ExcelVbaReference
    {
        /// <summary>
        /// Constructor.
        /// Defaults ReferenceRecordID to 0xD
        /// </summary>
        internal ExcelVbaReference()
        {
            ReferenceRecordID = 0xD;
        }
        /// <summary>
        /// The reference record ID. See MS-OVBA documentation for more info. 
        /// </summary>
        public int ReferenceRecordID { get; internal set; }
        /// <summary>
        /// The name of the reference
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// LibID
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string Libid { get; set; }
        /// <summary>
        /// A string representation of the object (the Name)
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Name;
        }
    }
}
