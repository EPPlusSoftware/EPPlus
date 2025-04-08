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
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// An exception thrown when the <see cref="ExcelPackage.License"/> hasn't been set.
    /// To set the license, use <seealso cref="EPPlusLicense.SetCommercial(string)"/>, <seealso cref="EPPlusLicense.SetNonCommercialOrganization(string)"/> or <seealso cref="EPPlusLicense.SetNonCommercialPersonal(string)"/>
    /// </summary>
    public class LicenseNotSetException : Exception
    {
        internal LicenseNotSetException(string message) : base(message)
        {

        }
    }
    /// <summary>
    /// An exception thrown when trying to set the obsolete property LicenseContext is set. <br/>
    /// Please use the <see cref="ExcelPackage.License"/> instead. <br/>
    /// To set the license, use <seealso cref="EPPlusLicense.SetCommercial(string)"/>, <seealso cref="EPPlusLicense.SetNonCommercialOrganization(string)"/> or <seealso cref="EPPlusLicense.SetNonCommercialPersonal(string)"/><br/>
    /// For more information see <seealso href ="http://epplussoftware.com/developers/licensenotsetexception"/>
    /// </summary>
    public class LicenseContextPropertyObsoleteException : Exception
    {
        internal LicenseContextPropertyObsoleteException(string message) : base(message)
        {

        }
    }
    /// <summary>
    /// An exception thrown when the license key cannot be validated.
    /// </summary>
    public class InvalidLicenseKeyException : Exception
    {
        internal InvalidLicenseKeyException(string message) : base(message)
        {

        }
    }
    /// <summary>
    /// An exception thrown when the license has expired for the version used.
    /// </summary>
    public class LicenseNotValidException : Exception
    {
        internal LicenseNotValidException(string message) : base(message)
        {

        }
    }
    /// <summary>
    /// An exception thrown when the license has expired for the version used.
    /// </summary>
    public class LicenseInformationException : Exception
    {
        internal LicenseInformationException(string message) : base(message)
        {

        }
    }
}

