using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using static OfficeOpenXml.EPPlusLicenseInfo;
namespace OfficeOpenXml
{
    internal class LicenseHandler
    {
        static readonly string _key = "<RSAKeyValue><Modulus>4iY/DJNamgOnXKWQCwuVqL7Jf5cAwZUWp23k6bjgUsiQNZzmDakeq2HRmlFTTUZrZC+frtAbxe6vFR9iqMvA9Q==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";
        static Uri licenseTextUri = new Uri("/EPPlusLicense.txt", UriKind.Relative);

        internal static void TagDocument(ExcelWorkbook wb)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            wb.Properties.Keywords = "EPPlus noncommercial use";
            if (string.IsNullOrEmpty(ExcelPackage.License.LegalName))
            {
                wb.Properties.Comments = $"This workbook has been created with EPPlus under The Polyform Noncommercial license: See https://polyformproject.org/licenses/noncommercial/1.0.0";
            }
            else
            {
                wb.Properties.Comments = $"This workbook has been created with EPPlus licensed to {ExcelPackage.License.LegalName} under The Polyform Noncommercial License: See https://polyformproject.org/licenses/noncommercial/1.0.0";
            }
            wb.Properties.Application = "EPPlus";
            wb.Properties.AppVersion = $"{version.Major}.{version.Minor}";

            if (ExcelPackage.License.LicenseType == EPPlusLicenseType.NonCommercialOrganization)
            {
                wb.Properties.Company = ExcelPackage.License.LegalName;
            }
            else
            {
                wb.Properties.Author = ExcelPackage.License.LegalName;
            }

            ZipPackagePart part;
            if (wb._package.ZipPackage.PartExists(licenseTextUri) == false)
            {
                part = wb._package.ZipPackage.CreatePart(licenseTextUri, "text/plain");
            }
            else
            {
                part = wb._package.ZipPackage.GetPart(licenseTextUri);
            }
            var stream = part.GetStream(FileMode.Create);
            var sw = new StreamWriter(stream);
            sw.WriteLine($"This workbook was created with the EPPlus library{(string.IsNullOrEmpty(ExcelPackage.License?.LegalName) ? "" : ", licensed to " + ExcelPackage.License?.LegalName)} under the Polyform Noncommercial license, see https://polyformproject.org/licenses/noncommercial/1.0.0");
            sw.WriteLine("For more information about EPPlus, see https://epplussoftware.com/");
            sw.Flush();

        }
        internal static void TrialTagDocument(ExcelWorkbook wb)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            wb.Properties.Keywords = "EPPlus Trial License "+ExcelPackage.License.LicenseInfo.LicenseNumber;
            wb.Properties.Comments = $"This workbook has been created with EPPlus using a trial license expiring: {ExcelPackage.License.LicenseInfo.LicenseValidTo:d}";
            wb.Properties.Application = "EPPlus";
            wb.Properties.AppVersion = $"{version.Major}.{version.Minor}";
            ZipPackagePart part;
            if (wb._package.ZipPackage.PartExists(licenseTextUri)==false)
            {
                part = wb._package.ZipPackage.CreatePart(licenseTextUri, "text/plain");
            }
            else
            {
                part = wb._package.ZipPackage.GetPart(licenseTextUri);
            }
            var stream = part.GetStream(FileMode.Create);
            var sw = new StreamWriter(stream);
            sw.WriteLine($"This workbook was created with the EPPlus library using a trial License: {ExcelPackage.License.LicenseInfo.LicenseNumber}.");
            sw.WriteLine("For more information about EPPlus, see https://epplussoftware.com/");
            sw.Flush();
        }

        private static DateTime GetValidToDate(DateTime toDate, EPPlusCommercialLicenseType licenseType, bool extendRenewal)
        {
            var validToDate = toDate;
            if(EnumUtil.HasNotFlag(licenseType, EPPlusCommercialLicenseType.Subscription))
            {
                return toDate;
            }
            if(toDate < DateTime.Today && EnumUtil.HasNotFlag(licenseType, EPPlusCommercialLicenseType.TemporaryKey))
            {
                validToDate = toDate.AddDays(30);
                if (extendRenewal)
                {
                    validToDate = validToDate.AddDays(15);
                }
            }
            return validToDate;
        }

        internal static bool ValidateLicenseKey(string licenseKey, bool extendRenewal, out EPPlusLicenseInfo licenseInfo, out string msg)
        {
            try
            {
                licenseKey = licenseKey.Trim('"').Trim();
                GetLicenseDataFromKey(licenseKey, out byte version, out string licenseNo, out DateTime fromDate, out DateTime toDate, out EPPlusCommercialLicenseType licenseType, out short numberOfLicenses, out byte[] signature, 512 / 8);
                var validToDate = GetValidToDate(toDate, licenseType, extendRenewal);
                licenseInfo = new EPPlusLicenseInfo()
                {
                    LicenseNumber = licenseNo,
                    LicenseType = licenseType,
                    LicenseValidFrom = fromDate,
                    LicenseValidTo = validToDate,
                    NumberOfLicensedDevelopers = numberOfLicenses
                };

                var tb = GetLicenseData(version, licenseNo, fromDate, toDate, (byte)licenseType, numberOfLicenses);
                try
                {
#if NET35
                    var rsaClient = new RSACryptoServiceProvider();
#else
                    var rsaClient = RSA.Create();
#endif

                    var rsa2 = RSA.Create();
                    rsaClient.FromXmlString(_key);
#if NET35
                    if (rsaClient.VerifyData(tb, "2.16.840.1.101.3.4.2.1", signature))
#else
                    if (rsaClient.VerifyData(tb, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
#endif
                    {
                        return ValidateLicenseDates(licenseInfo, out msg);
                    }
                    else
                    {
                        licenseInfo.Status = EPPlusLicenseStatus.InvalidLicenseKey;
                        msg = "The license key is not valid. Please use the license key as stated on your license document or as displayed on your account at https://epplussoftware.com";
                        return false;
                    }
                }
                catch (PlatformNotSupportedException)
                {
                    return ValidateLicenseDates(licenseInfo, out msg);
                }
            }
            catch
            {
                licenseInfo = new EPPlusLicenseInfo()
                {
                    Status = EPPlusLicenseStatus.InvalidLicenseKey
                };
                msg="The license key is not in a valid format. Please use the license key as stated on your license document or as displayed on your account at https://epplussoftware.com";
                return false;
            }
        }

        private static bool ValidateLicenseDates(EPPlusLicenseInfo licenseInfo, out string msg)
        {
            if (licenseInfo.LicenseValidFrom > DateTime.Today)
            {
                licenseInfo.Status = EPPlusLicenseStatus.NotValidForThisVersion;
                msg = $"This EPPlus license is not valid until {licenseInfo.LicenseValidFrom:d}.";
                return false;
            }
            var vd = DateTime.Parse(EPPlusLicense._versionDate, CultureInfo.InvariantCulture);
            if (EnumUtil.HasFlag(licenseInfo.LicenseType, EPPlusCommercialLicenseType.Subscription))
            {
                DateTime bd = vd;
                if(Debugger.IsAttached)
                {
                    bd = DateTime.Today; 
                }

                var extendUnderRenewal = ExcelPackage.License.ExtendUnderRenewal;
                if (licenseInfo.LicenseValidTo.AddDays(extendUnderRenewal ? 15 : 0) < bd)
                {
                    msg = $"This EPPlus license key is no longer valid, it expired on {licenseInfo.LicenseValidTo:d}. If your license has been renewed, please use the new license key found in your license document or in your account at https://epplussoftware.com";
                    if(extendUnderRenewal==false && licenseInfo.LicenseValidTo.AddDays(15) > bd)
                    {
                        msg += " To get 15 additional days validity off this key, you can set the ExcelPackage.License.ExtendUnderRenewal to true.";
                    }
                    licenseInfo.Status = EPPlusLicenseStatus.Expired;
                    return false;
                }
            }
            else
            {
                if (licenseInfo.LicenseValidTo < vd)
                {
                    licenseInfo.Status = EPPlusLicenseStatus.NotValidForThisVersion;
                    msg = $"This license key is not valid for EPPlus versions released after {licenseInfo.LicenseValidTo:d}. This EPPlus version was released on {EPPlusLicense._versionDate:d}. If your license has been renewed, please use the new license key found in your license document or in your account at https://epplussoftware.com";
                    return false;
                }
            }
            msg = null;
            licenseInfo.Status = EPPlusLicenseStatus.Valid;
            return true;
        }

        static byte[] GetLicenseData(byte version, string licenseNo, DateTime fromDate, DateTime toDate, byte licenseType, short numberOfLicenses)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                var br = new BinaryWriter(ms);
                br.Write(version);
                var licenseNoBytes = ASCIIEncoding.ASCII.GetBytes(licenseNo);
                br.Write((byte)licenseNoBytes.Length);
                br.Write(licenseNoBytes);
                var baseDate = new DateTime(fromDate.Year, 1, 1);
                br.Write((short)fromDate.Year);
                br.Write((short)(fromDate - baseDate).Days);
                br.Write((short)(toDate - baseDate).Days);
                br.Write(licenseType);
                br.Write(numberOfLicenses);
                br.Flush();
                var tb = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(tb, 0, (int)ms.Length);
                return tb;
            }
        }
        internal static void GetLicenseDataFromKey(string lk, out byte version, out string licenseNo, out DateTime fromDate, out DateTime toDate, out EPPlusCommercialLicenseType licenseType, out short noOfLicenses, out byte[] signature, int size)
        {        
            var by = Convert.FromBase64String(lk);
            using (var ms = RecyclableMemory.GetStream(by))
            {
                var br = new BinaryReader(ms);
                signature = br.ReadBytes(size);
                version = br.ReadByte();
                var len = (int)br.ReadByte();
                licenseNo = ASCIIEncoding.ASCII.GetString(br.ReadBytes(len));
                var yearBase = br.ReadInt16();
                var baseYear = new DateTime(yearBase, 1, 1);
                var fdDays = br.ReadInt16();
                fromDate = baseYear.AddDays(fdDays);
                var tdDays = br.ReadInt16();
                toDate = baseYear.AddDays(tdDays);
                licenseType = (EPPlusCommercialLicenseType)br.ReadByte();
                noOfLicenses = br.ReadInt16();
            }
        }

        internal static void ApplyCommercialLicense(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(licenseTextUri))
            {
                wb._package.ZipPackage.DeletePart(licenseTextUri);
            }
        }
    }
}
