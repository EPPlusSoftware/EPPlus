using OfficeOpenXml.Packaging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Drawing.Vml;
using System.Linq;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ExcelDigitalSignatureCollection : IEnumerable<ExcelDigitalSignature>
    {
        ExcelPackage _package;
        ExcelWorkbook _wb;
        Uri _sigOrigin;
        XmlNamespaceManager _ns;

        List<ExcelDigitalSignature> _signatures;
        List<DigitalSignatureLine> _signatureLines = new List<DigitalSignatureLine>();

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

        public ExcelDigitalSignature this[int index]
        {
            get { return _signatures[index]; }
            set { _signatures[index] = value; }
        }

        private void LoadSignatures()
        {
            var originPart = _package.ZipPackage.GetPart(_sigOrigin);
            var rels = originPart.GetRelationships();

            int i = 1;

            foreach(var rel in rels)
            {
                var adjustedUri = new Uri("_xmlsignatures/" + rel.TargetUri.OriginalString, UriKind.Relative);
                var part = _package.ZipPackage.GetPart(adjustedUri);
                ReadPartXml(part, i);
                i++;
            }
        }

        private void ReadPartXml(ZipPackagePart part, int num)
        {
            var signatureXml = new XmlDocument();
            signatureXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;

            var digitalSignature = new ExcelDigitalSignature(_wb, _ns, part, num);
            _signatures.Add(digitalSignature);
        }

        public ExcelDigitalSignature AddSignature(X509Certificate2 certificate, CommitmentType cType = CommitmentType.None, string purposeForSigning = "")
        {
            var digSig = new ExcelDigitalSignature(_wb, _ns, _signatures.Count + 1);

            digSig.Certificate = certificate;
            digSig.commitmentType = cType;
            digSig.PurposeForSigning = purposeForSigning;

            _signatures.Add(digSig);
            return digSig;
        }

        public DigitalSignatureLine AddSignatureLine(X509Certificate2 certificate, ExcelWorksheet ws, CommitmentType cType = CommitmentType.None, string purposeForSigning = "")
        {
            _signatureLines.Add(new DigitalSignatureLine(ws));
            return _signatureLines.Last();
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
