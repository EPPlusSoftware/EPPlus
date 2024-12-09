using OfficeOpenXml.VBA;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Hash algorithm used for digital signatures.
    /// </summary>
    public enum DigitalSignatureHashAlgorithm
    {
        /// <summary>
        /// Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA1 = VbaSignatureHashAlgorithm.SHA1,
        /// <summary>
        /// Specifies that the SHA-256 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA256 = VbaSignatureHashAlgorithm.SHA256,
        /// <summary>
        /// Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA384 = VbaSignatureHashAlgorithm.SHA384,
        /// <summary>
        /// Specifies that the SHA-512 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA512 = VbaSignatureHashAlgorithm.SHA512,
    }
}
