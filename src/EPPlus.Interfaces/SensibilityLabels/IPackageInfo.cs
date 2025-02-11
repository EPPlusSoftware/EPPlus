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
using System.IO;

namespace OfficeOpenXml.Interfaces.SensitivityLabels;

/// <summary>
/// An interface for passing information of decryption/encryption of a protected package.
/// </summary>
public interface IPackageInfo
{
    /// <summary>
    /// The unencrypted package stream. 
    /// </summary>
    public MemoryStream PackageStream { get; set; }
    /// <summary>
    /// Information passed from the <see cref="DecryptPackageAsync" />  to the <see cref="ApplyLabelAndSavePackageAsync" />.
    /// Use this property to pass protection information that should be retained when saving the package. 
    /// This property will be set to null, if a new Sensibilty label is applied to the package.
    /// </summary>
    public object ProtectionInformation { get; set; }
    /// <summary>
    /// If a sensibility label has been set using EPPlus, this property contains the id of the new label.
    /// </summary>
    public string ActiveLabelId { get; set; }
}
#endif