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
namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// Specifies whether the target is inside or outside the System.IO.Packaging.Package.
    /// </summary>
    public enum TargetMode
    {
        /// <summary>
        /// The relationship references a part that is inside the package.
        /// </summary>
        Internal = 0,
        /// <summary>
        /// The relationship references a resource that is external to the package.
        /// </summary>
        External = 1,
    }
}