using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    enum PartType
    {
        Part,
        RelPart
    }

    internal struct PartWithXml()
    {
        internal string UriKey;
        internal string Xml;
        internal PartType PartType;
    }

    internal class ZipPackageXmlManifest
    {
        internal List<PartWithXml> partXmlList;

        internal ZipPackageXmlManifest()
        {
            partXmlList = new List<PartWithXml>();
        }

        internal void AddPart(string uriKey, string xml, PartType partType)
        {
            var newPart = new PartWithXml { UriKey = uriKey, Xml = xml, PartType = partType};
            partXmlList.Add(newPart);
        }
    }
}
