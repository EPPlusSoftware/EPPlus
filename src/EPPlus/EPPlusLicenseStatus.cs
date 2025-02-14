namespace OfficeOpenXml
{
    /// <summary>
    /// Status of a commercial license
    /// </summary>
    public enum EPPlusLicenseStatus
    {
        /// <summary>
        /// The license key is valid for this version.
        /// </summary>
        Valid,
        /// <summary>
        /// The license key is invalid.
        /// </summary>
        InvalidLicenseKey,
        /// <summary>
        /// The subscription license has expired.
        /// </summary>
        Expired,
        /// <summary>
        /// The perpetual license is valid, but requires a renewal for this version.
        /// </summary>
        NotValidForThisVersion
    }
}
