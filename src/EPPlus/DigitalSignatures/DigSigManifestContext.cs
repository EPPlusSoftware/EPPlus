using OfficeOpenXml.Utils.EncodingUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    enum ePartType
    {
        Part,
        RelPart
    }

    //Saves hashed version of parts and xml of rels
    internal class PartWithXml()
    {
        internal string UriKey;
        internal ePartType PartType;
        internal string StringData;
    }

    internal class DigSigManifestContext
    {
        internal List<PartWithXml> partXmlList;
        HashAlgorithm _hashAlgorithm;

        internal DigSigManifestContext(DigitalSignatureHashAlgorithm hashAlgorithm)
        {
            SetHashAlgorithm(hashAlgorithm);
            partXmlList = new List<PartWithXml>();
        }

        private void SetHashAlgorithm(DigitalSignatureHashAlgorithm hashAlgorithm)
        {
            _hashAlgorithm = EncodeUtil.GetHashProvider(hashAlgorithm);
        }

        internal void AddPart(string uriKey, ePartType partType, byte[] bytes, int size)
        {
            var newPart = new PartWithXml { UriKey = uriKey, PartType = partType, StringData = Encoding.UTF8.GetString(bytes,0, size) };
            partXmlList.Add(newPart);
        }

        internal void AddPartHashOnly(string uriKey, ePartType partType, byte[] bytes, int size)
        {
            var hashResult = _hashAlgorithm.ComputeHash(bytes, 0, size);
            var hashConvert = Convert.ToBase64String(hashResult);
            var newPart = new PartWithXml { UriKey = uriKey, PartType = partType, StringData = hashConvert };
            partXmlList.Add(newPart);
        }
    }
}
