using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.IO;
using static System.Net.WebRequestMethods;
using System.Reflection;
namespace OfficeOpenXml
{
    internal class LicenseHandler
    {
        static readonly string _key = "<RSAKeyValue><Modulus>vKJxhqMkmgoCZFBU4/RWfQ86PaNA2Adj3ZbhmN7Op3YIJNy+YhduR9/nm4ynM2XduXlFFZ6xNQgKl3xqgm9pcQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        internal static void TagDocument(ExcelWorkbook wb)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            wb.Properties.Keywords = "EPPlus Non-Commercial Use";
            wb.Properties.Comments = "This workbook has been created with EPPlus licensed under The Polyform Non-Commercial License: See https://polyformproject.org/licenses/noncommercial/1.0.0";
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

            var part = wb._package.ZipPackage.CreatePart(new Uri("/EPPlusLicense.txt", UriKind.Relative), "text/plain");
            var stream = part.GetStream();            
            var sw=new StreamWriter(stream);
            sw.WriteLine("This workbook was created by the EPPlus library licensed under the Polyform Non-Commercial license, see https://polyformproject.org/licenses/noncommercial/1.0.0");
            sw.WriteLine("For more information about EPPlus, see https://epplussoftware.com/");
            sw.Flush();            
        }

        internal static bool ValidateLicenseKey(string licenseKey, EPPlusLicenseInfo licenseInfo=null)
        {
            GetLicenseDataFromKey(licenseKey, out byte version, out string licenseNo, out DateTime fromDate, out DateTime toDate, out byte licenseType, out short numberOfLicenses, out byte[] signature, 512 / 8);
            if (licenseInfo != null)
            {
                licenseInfo.LicenseNumber = licenseNo;
                licenseInfo.LicenseType = licenseType;
                licenseInfo.LicenseValidFrom = fromDate;
                licenseInfo.LicenseValidTo = toDate;
                licenseInfo.NumberOfLicenses = numberOfLicenses;
            }
            var tb = GetLicenseData(version, licenseNo, fromDate, toDate, licenseType, numberOfLicenses); ;
            var rsaClient = new RSACryptoServiceProvider();
            rsaClient.FromXmlString(_key);
            var oid = CryptoConfig.MapNameToOID("SHA256");
            if (rsaClient.VerifyData(tb, oid, signature))
            {                
                return true;
            }
            else
            {
                throw new InvalidLicenseKeyException("The license key is not valid. Please use the license key as stated on your license document");
            }
        }
        static byte[] GetLicenseData(byte version, string licenseNo, DateTime fromDate, DateTime toDate, byte licenseType, short numberOfLicenses)
        {
            var ms = new MemoryStream();
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
        internal static void GetLicenseDataFromKey(string lk, out byte version, out string licenseNo, out DateTime fromDate, out DateTime toDate, out byte licenseType, out short noOfLicenses, out byte[] signature, int size)
        {
            var by = Convert.FromBase64String(lk);
            var br = new BinaryReader(new MemoryStream(by));
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
            licenseType = br.ReadByte();
            noOfLicenses = br.ReadInt16();
        }
    }
}
