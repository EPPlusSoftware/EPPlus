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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Encryption Algorithm
    /// </summary>
    public enum EncryptionAlgorithm
    {
        /// <summary>
        /// 128-bit AES. Default
        /// </summary>
        AES128,
        /// <summary>
        /// 192-bit AES.
        /// </summary>
        AES192,
        /// <summary>
        /// 256-bit AES. 
        /// </summary>
        AES256
    }
    /// <summary>
    /// The major version of the Encryption 
    /// </summary>
    public enum EncryptionVersion
    {
        /// <summary>
        /// Standard Encryption.
        /// Used in Excel 2007 and previous version with compatibility pack.
        /// <remarks>Default AES 128 with SHA-1 as the hash algorithm. Spincount is hardcoded to 50000</remarks>
        /// </summary>
        Standard,
        /// <summary>
        /// Agile Encryption.
        /// Used in Excel 2010-
        /// Default.
        /// </summary>
        Agile,
        /// <summary>
        /// The workbook is protected by a sensitiviy label.
        /// For EPPlus to work with this type of encryption you need to set a <see cref="OfficeOpenXml.SensitivityLabels.ExcelSensibilityLabels.SensibilityLabelHandler"/> that handels the decryption/encryption using the Microsoft MIPS API.
        /// </summary>
        ProtectedBySensibilityLabel
    }
    /// <summary>
    /// How and if the workbook is encrypted
    ///<seealso cref="ExcelProtection"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    /// </summary>
    public class ExcelEncryption
    {
        /// <summary>
        /// Constructor
        /// <remarks>Default AES 256 with SHA-512 as the hash algorithm. Spincount is set to 100000</remarks>
        /// </summary>
        internal ExcelEncryption()
        {
            Algorithm = EncryptionAlgorithm.AES256;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="encryptionAlgorithm">Algorithm used to encrypt the package. Default is AES128</param>
        internal ExcelEncryption(EncryptionAlgorithm encryptionAlgorithm)
        {
            Algorithm = encryptionAlgorithm;
        }
        bool _isEncrypted = false;
        /// <summary>
        /// Is the package encrypted
        /// </summary>
        public bool IsEncrypted
        {
            get
            {
                return _isEncrypted;
            }
            set
            {
                _isEncrypted = value;
                if (_isEncrypted)
                {
                    if (_password == null) _password = "";
                }
                else
                {
                    _password = null;
                }
            }
        }
        string _password = null;
        /// <summary>
        /// The password used to encrypt the workbook.
        /// </summary>
        public string Password
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
                _isEncrypted = (value != null);
            }
        }
        /// <summary>
        /// Algorithm used for encrypting the package. Default is AES 128-bit for standard and AES 256 for agile
        /// </summary>
        public EncryptionAlgorithm Algorithm { get; set; }
        private EncryptionVersion _version = EncryptionVersion.Agile;
        /// <summary>
        /// The version of the encryption.        
        /// </summary>
        public EncryptionVersion Version
        {
            get
            {
                return _version;
            }
            set
            {
                if (value != Version)
                {
                    if (value == EncryptionVersion.Agile)
                    {
                        Algorithm = EncryptionAlgorithm.AES256;
                    }
                    else
                    {
                        Algorithm = EncryptionAlgorithm.AES128;
                    }
                    _version = value;
                }
            }
        }
        /// <summary>
        /// Encrypts a stream using the office encryption.
        /// </summary>
        /// <param name="stream">The stream containing the non-encrypted package.</param>
        /// <param name="password">The password to encrypt with</param>
        /// <param name="encryptionVersion">The encryption version</param>
        /// <param name="algorithm">The algorithm to use for the encryption</param>
        /// <returns>A MemoryStream containing the encypted package</returns>
        public static MemoryStream EncryptPackage(Stream stream, string password, EncryptionVersion encryptionVersion=EncryptionVersion.Agile, EncryptionAlgorithm algorithm = EncryptionAlgorithm.AES256)
        {
            var e = new Encryption.EncryptedPackageHandler(null);
            if(stream.CanRead==false)
            {
                throw new InvalidOperationException("Stream must be readable");
            }
            if (stream.CanSeek)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
            
            var b = new byte[stream.Length];
            stream.Read(b, 0, (int)stream.Length);
            return e.EncryptPackage(b, new ExcelEncryption { Password = password, Algorithm = algorithm, Version = encryptionVersion });
        }
        /// <summary>
        /// Decrypts a stream using the office encryption.
        /// </summary>
        /// <param name="stream">The stream containing the encrypted package.</param>
        /// <param name="password">The password to decrypt with</param>
        /// <returns>A memorystream with the encypted package</returns>
        public static MemoryStream DecryptPackage(Stream stream, string password)
        {
            var e = new Encryption.EncryptedPackageHandler(null);
            if(stream==null)
            {
                throw new ArgumentNullException("Stream must not be null");
            }
            if (stream.CanRead == false)
            {
                throw new InvalidOperationException("Stream must be readable");
            }
            if (stream.CanSeek)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
#if (NET35)
            else
            {
                throw new InvalidOperationException("Stream must be seekable");
            }
#endif

            MemoryStream ms;
            if(stream is MemoryStream)
            {
                ms = (MemoryStream)stream;
            }
            else
            {
#if (NET35)
                var b = new byte[stream.Length];
                stream.Read(b, 0, (int)stream.Length);
                ms = new MemoryStream(b);
#else
                ms = RecyclableMemory.GetStream();
                stream.CopyTo(ms);                
#endif
            }
            return e.DecryptPackage(ms, new ExcelEncryption() { Password = password, _isEncrypted=true });
        }

    }
}
