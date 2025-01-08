using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// The type of commercial license
    /// </summary>
    [Flags]
    public enum EPPlusCommercialLicenseType
    {
        /// <summary>
        /// A subscription license.
        /// </summary>
        Subscription = 0x1,
        /// <summary>
        /// A perpetual license.
        /// </summary>
        Perpetual = 0x2,
        /// <summary>
        /// A perpetual license package.
        /// </summary>
        Package = 0x4,
        /// <summary>
        /// A custom license, for example an enterprise or site license.
        /// </summary>
        Custom = 0x8,
        /// <summary>
        /// A pay-as-you go license. This license can no longer be purchased.
        /// </summary>
        PayAsYouGo = 0x10,
        /// <summary>
        /// This license is granted for trial purposes, normaly for 32 days.
        /// </summary>
        Trial = 0x20,
        /// <summary>
        /// This license has been temporary extended to complete the renewal process.
        /// </summary>
        TemporaryKey = 0x40
    }
    public enum EPPlusLicenseStatus
    {
        /// <summary>
        /// The license key is valid for this version.
        /// </summary>
        IsValid,
        /// <summary>
        /// The license key is invalid.
        /// </summary>
        IsInvalidLicenseKey,
        /// <summary>
        /// The subscription license has expired.
        /// </summary>
        IsExpired,
        /// <summary>
        /// The perpetual license is valid, but requires a renewal for this version.
        /// </summary>
        IsNotValidForThisVersion
    }
}
