using OfficeOpenXml.Packaging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.Security.Cryptography.X509Certificates;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Collection of digital signatures
    /// </summary>
    public class ExcelDigitalSignatureCollection : IEnumerable<ExcelDigitalSignature>
    {
        ExcelPackage _package;
        ExcelWorkbook _wb;
        Uri _sigOrigin;
        XmlNamespaceManager _ns;

        List<ExcelDigitalSignature> _signatures;

        internal ExcelDigitalSignatureCollection(ExcelWorkbook wb, XmlNamespaceManager ns)
        {
            _package = wb._package;
            _wb = wb;
            _ns = ns;
            _signatures = new List<ExcelDigitalSignature>();
        }

        internal ExcelDigitalSignatureCollection(ExcelWorkbook wb, XmlNamespaceManager ns, Uri signatureOriginUri)
        {
            _package = wb._package;
            _wb = wb;
            _sigOrigin = signatureOriginUri;
            _ns = ns;

            _signatures = new List<ExcelDigitalSignature>();
            LoadSignatures();
        }
        IEnumerator<ExcelDigitalSignature> IEnumerable<ExcelDigitalSignature>.GetEnumerator()
        {
            for (int i = 0; i < _signatures.Count; i++)
            {
                yield return _signatures[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _signatures.GetEnumerator();
        }

        /// <summary>
        /// Get the signature at index.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelDigitalSignature this[int index]
        {
            get { return _signatures[index]; }
            set { _signatures[index] = value; }
        }

        private void LoadSignatures()
        {
            var originPart = _package.ZipPackage.GetPart(_sigOrigin);
            var rels = originPart.GetRelationships();

            foreach(var rel in rels)
            {
                var adjustedUri = new Uri("_xmlsignatures/" + rel.TargetUri.OriginalString, UriKind.Relative);
                var part = _package.ZipPackage.GetPart(adjustedUri);
                ReadPartXml(part);
            }
        }

        private void ReadPartXml(ZipPackagePart part)
        {
            var signatureXml = new XmlDocument();
            signatureXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;

            var digitalSignature = new ExcelDigitalSignature(_wb, _ns, part);

            _signatures.Add(digitalSignature);
        }

        /// <summary>
        /// Add digital signature
        /// Requires a valid X509Certificate2 with a private key.
        /// </summary>
        /// <param name="certificate"></param>
        /// <returns></returns>
        public ExcelDigitalSignature Add(X509Certificate2 certificate)
        {
            var digSig = new ExcelDigitalSignature(_wb, _ns, _signatures.Count + 1);

            digSig.Certificate = certificate;

            _signatures.Add(digSig);
            return digSig;
        }

        /// <summary>
        /// Remove digital signature
        /// </summary>
        /// <param name="signature"></param>
        public void Remove(ExcelDigitalSignature signature)
        {
            _wb._package.ZipPackage.DeletePart(new Uri(signature.PartUri, UriKind.Relative));

            _signatures.Remove(signature);
        }

        internal ExcelDigitalSignature GetSignatureBySignatureLineGuid(Guid id)
        {
            foreach (var sig in _signatures)
            {
                if (sig.SignatureLine != null && sig.SignatureLine.SetupID.Equals(id))
                {
                    return sig;
                }
            }
            return null;
        }
        internal ExcelDigitalSignature GetSignatureByFileName(string fileName)
        {
            foreach (var sig in _signatures)
            {
                if(sig.PartUri == fileName)
                {
                    return sig;
                }
            }
            return null;
        }
    }
}
