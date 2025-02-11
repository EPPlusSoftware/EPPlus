/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial licenseXml to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/29/2024         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Encryption
{
    internal class DataSpaceReference
    {
        internal enum eReferenceType
        {
            Stream=0,
            Storage=1
        }
        public eReferenceType ReferenceType { get; set; }
        public string Name { get; set; }
    }
}
