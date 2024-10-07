namespace OfficeOpenXml
{
    public enum EPPlusLicenseSource
    {
        /// <summary>
        /// The license is not set.
        /// </summary>
        NotSet,
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