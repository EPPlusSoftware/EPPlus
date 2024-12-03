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

    internal struct PartWithXml()
    {
        internal string UriKey;
        internal string Xml;
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

        internal void AddPart(string uriKey, string xml, ePartType partType, byte[] bytes)
        {
            var newPart = new PartWithXml { UriKey = uriKey, Xml = xml, PartType = partType, Bytes = bytes};
            partXmlList.Add(newPart);
        }
    }
}
