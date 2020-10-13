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
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Security.Cryptography;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// File sharing settings for the workbook.
    /// </summary>
    public class ExcelFileSharing : XmlHelper
    {
        internal ExcelFileSharing(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// Set the workbook to readonly for anyone
        /// </summary>
        /// <param name="userName">The name of the person enforcing the writeprotection</param>
        /// <param name="password">The password. Must not be empty.</param>
        public void SetReadOnly(string userName, string password)
        {
            if(string.IsNullOrEmpty(password.Trim()))
            {
                throw new ArgumentOutOfRangeException("password", "Password must not be null or empty");
            }
            UserName = userName;
            HashAlogorithm = eHashAlogorithm.SHA512;

            var s = new byte[16];
            var rnd = RandomNumberGenerator.Create();
            rnd.GetBytes(s);
            SaltValue = s;
            SpinCount = 100000;

            HashValue = EncryptedPackageHandler.GetPasswordHashSpinAppending(SHA512.Create(), SaltValue, password, SpinCount, 64);
        }
        /// <summary>
        /// Remove any write protection set on the workbook
        /// </summary>
        public void RemoveReadOnly()
        {
            DeleteNode("d:fileSharing");
        }
        internal eHashAlogorithm HashAlogorithm
        {
            get
            {
                return GetHashAlogorithm(GetXmlNodeString("d:fileSharing/@algorithmName"));
            }
            private set
            {
                SetXmlNodeString("d:fileSharing/@algorithmName", SetHashAlogorithm(value));
            }
        }

        private string SetHashAlogorithm(eHashAlogorithm value)
        {
            switch(value)
            {
                case eHashAlogorithm.SHA512:
                    return "SHA-512";
                default:
                    throw new NotSupportedException("EPPlus only support SHA 512 hashing for file sharing");
            }
        }

        private eHashAlogorithm GetHashAlogorithm(string v)
        {
            switch (v)
            {
                case "RIPEMD-128":
                    return eHashAlogorithm.RIPEMD128;
                case "RIPEMD-160":
                    return eHashAlogorithm.RIPEMD160;
                case "SHA-1":
                    return eHashAlogorithm.SHA1;
                case "SHA-256":
                    return eHashAlogorithm.SHA256;
                case "SHA-384":
                    return eHashAlogorithm.SHA384;
                case "SHA-512":
                    return eHashAlogorithm.SHA512;
                default:
                    return v.ToEnum(eHashAlogorithm.SHA512);
            }
        }

        internal int SpinCount
        {
            get
            {
                return GetXmlNodeInt("d:fileSharing/@spinCount");
            }
            set
            {
                SetXmlNodeInt("d:fileSharing/@spinCount", value);
            }
        }
        internal byte[] SaltValue
        {
            get
            {
                var s = GetXmlNodeString("d:fileSharing/@saltValue");
                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }
                return null;
            }
            set
            {
                SetXmlNodeString("d:fileSharing/@saltValue", Convert.ToBase64String(value));
            }
        }
        internal byte[] HashValue
        {
            get
            {
                var s = GetXmlNodeString("d:fileSharing/@hashValue");
                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }
                return null;
            }
            set
            {
                SetXmlNodeString("d:fileSharing/@hashValue", Convert.ToBase64String(value));
            }
        }
        /// <summary>
        /// whether the application alerts the user that the file be marked as read-only
        /// </summary>
        public bool IsReadOnly
        {
            get
            {
                return ExistNode("d:fileSharing/@hashValue");
            }
        }
        /// <summary>
        /// The name of the person enforcing the write protection.
        /// </summary>
        public string UserName
        {
            get
            {
                return GetXmlNodeString("d:fileSharing/@userName");
            }
            set
            {
                SetXmlNodeString("d:fileSharing/@userName", value);
            }
        }
        /// <summary>
        /// If opening the workbook in readonly is the recommended by the author.
        /// </summary>
        public bool ReadOnlyRecommended
        {
            get
            {
                return GetXmlNodeBool("d:fileSharing/@readOnlyRecommended");
            }
            set
            {
                if(IsReadOnly==false)
                {
                    throw new InvalidOperationException("Can only set this property when workbook is readonly.");
                }

                SetXmlNodeBool("d:fileSharing/@readOnlyRecommended", value);
            }
        }
    }
}