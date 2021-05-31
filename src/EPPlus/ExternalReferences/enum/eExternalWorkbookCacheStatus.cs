/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/26/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// The status of an external workbooks cache.
    /// </summary>
    public enum eExternalWorkbookCacheStatus
    {
        /// <summary>
        /// Cache has not been updated. Saving an external reference with this status will update the cache on save.
        /// </summary>
        NotUpdated,
        /// <summary>
        /// Cache has been loaded from the external reference cache within the package.
        /// </summary>
        LoadedFromPackage,
        /// <summary>
        /// Update of the cache failed. Any loaded data from the package is still available. 
        /// </summary>
        Failed,
        /// <summary>
        /// The cache has been successfully updated
        /// </summary>
        Updated
    }
}
