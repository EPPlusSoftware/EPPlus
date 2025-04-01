namespace OfficeOpenXml
{
    /// <summary>
    /// The source method how a license is set.
    /// </summary>
    public enum EPPlusLicenseSource
    {
        /// <summary>
        /// The license is set via the methods in <see cref="EPPlusLicense"/>
        /// </summary>
        Code,
        /// <summary>
        /// The license is set via an user or process environment variable.
        /// </summary>
        EnvironmentVariable,
        /// <summary>
        /// The license is set via the application configuration file
        /// </summary>
        ConfigFile
    }
}