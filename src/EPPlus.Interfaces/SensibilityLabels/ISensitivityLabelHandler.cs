/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/18/2024         EPPlus Software AB       EPPlus 8
 *************************************************************************************************/
#if(!NET35)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeOpenXml.Interfaces.SensitivityLabels;
/// <summary>
/// An interface that should be applied to a handler that integrates into EPPlus to handle sensibility labels. 
/// The handler should apply or remove sensibility labels, decrypt/encryt protected packages and add extended information about the sensibility labels.
/// </summary>
public interface ISensitivityLabelHandler
{
    /// <summary>
    /// Called to initiate the handler, when it is set.
    /// </summary>
    /// <returns></returns>
    public Task InitAsync();
    /// <summary>
    /// Called to decrypt a protected package.
    /// </summary>
    /// <param name="packageStream">The encrypted package stream</param>
    /// <param name="Id">The unique id of the package.</param>
    /// <returns>The decrypted package with any neccessary protection information.</returns>
    public Task<IPackageInfo> DecryptPackageAsync(MemoryStream packageStream, string Id);
    /// <summary>
    /// Called when the package is saved in EPPlus and should apply the active label using the Microsoft Information Protection SDK
    /// </summary>
    /// <param name="package">The package stream</param>
    /// <param name="Id">The unique package Id</param>
    /// <returns></returns>
    public Task<MemoryStream> ApplyLabelAndSavePackageAsync(IPackageInfo package, string Id);
    /// <summary>
    /// Should update the supplied list of sensibility lables with name, description and other properties not present in the Sensibility Label XML document inside the package.
    /// </summary>
    /// <param name="list">The list of lables to update.</param>
    /// <param name="Id">The unique id of the package.</param>
    public void UpdateLabelList(IEnumerable<IExcelSensibilityLabel> list, string Id);
    /// <summary>
    /// Get all labels from the MIPS api.
    /// </summary>
    /// <param name="Id"></param>
    /// <returns></returns>
    public IEnumerable<IExcelSensibilityLabel> GetLabels(string Id);
}
#endif