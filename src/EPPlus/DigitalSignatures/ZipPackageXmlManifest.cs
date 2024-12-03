using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    enum ePartType
    {
        Part,
        RelPart
    }

    internal class PartWithXml()
    {
        internal string UriKey;
        internal ePartType PartType;
        internal byte[] Bytes;
    }

    internal class ZipPackageXmlManifest
    {
        internal List<PartWithXml> partXmlList;

        internal ZipPackageXmlManifest()
        {
            partXmlList = new List<PartWithXml>();
        }

        internal void AddPart(string uriKey, ePartType partType, byte[] bytes)
        {
            var newPart = new PartWithXml { UriKey = uriKey, PartType = partType, Bytes = bytes};
            partXmlList.Add(newPart);
        }
    }
}
