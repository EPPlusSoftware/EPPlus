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
#if (!NET35)
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core;
using OfficeOpenXml.Encryption;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces;
using System;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.SensitivityLabels
{
    public class ExcelSensibilityLabels
    {
        /// <summary>
        /// If you want your workbooks to be marked with sensibility lables, you can add a handler for authentication, encryption and decryption using the Microsoft Information Protection SDK.
        /// For more information
        /// </summary>
        public static ISensitivityLabelHandler SensibilityLabelHandler { get; set; }

        ExcelPackage _pck;
        XmlHelper _xmlHelper;
        XmlNamespaceManager _nsm;
        internal ExcelSensibilityLabels(ExcelPackage pck)
        {
            _pck = pck;
            InitNsm();
            Labels = new EPPlusReadOnlyList<IExcelSensibilityLabel>();
            LoadLabelsFromPart();
        }

        internal ExcelSensibilityLabels(ExcelPackage pck, SensibilityLabelInfo si)
        {
            _pck = pck;
            Labels = new EPPlusReadOnlyList<IExcelSensibilityLabel>();
            InitNsm();
            var xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, si.LabelXml, Encoding.UTF8);
            LoadLabelsFromXmlDocument(xml);
        }
        private void InitNsm()
        {
            NameTable nt = new NameTable();
            _nsm = new XmlNamespaceManager(nt);
            _nsm.AddNamespace("clbl", ExcelPackage.schemaMipLabelMetadata);
        }
        private void LoadLabelsFromPart()
        {
            var part = _pck.ZipPackage.GetPartByContentType(ContentTypes.contentTypeClassificationLabels);
            if (part == null) return;
            var xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, part.GetStream());

            LoadLabelsFromXmlDocument(xml);
        }

        private void LoadLabelsFromXmlDocument(XmlDocument xml)
        {
            foreach (XmlElement element in xml.SelectNodes("/clbl:labelList/clbl:label", _nsm))
            {
                Labels.Add(ExcelSensibilityLabel.CreateFromElement(_nsm, element));
            }
        }
        
        public void SetActiveLabel(string name)
        {
            if(SensibilityLabelHandler==null)
            {
                throw (new MissingSensibilityHandlerException("No sensibility label handler is set. Please set the property ExcelSensibilityLabels.SensibilityLabelHandler"));
            }


            SensibilityLabelHandler.SetActiveLabel(name);
        }

        public EPPlusReadOnlyList<IExcelSensibilityLabel> Labels
        {
            get;
        }
    }

    public class ExcelSensibilityLabel : IExcelSensibilityLabel
    {
        /// <summary>
        /// The sensitivity label id. Guid.
        /// </summary>
        public string Id { get; internal set; }
        /// <summary>
        /// The name of the sensibility label. If no <see cref="ExcelSensibilityLabels.SensibilityLabelHandler"/> is set this property will always be empty.
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// The description of the sensibility label. If no <see cref="ExcelSensibilityLabels.SensibilityLabelHandler"/> is set this property will always be empty.
        /// </summary>
        public string Description { get; internal set; }
        /// <summary>
        /// If the sensibility label is enabled. Only one sensibility label can be enabled in the list.
        /// </summary>
        public bool Enabled { get; internal set; }
        /// <summary>
        /// If the sensibility label is removed. If the sensibility label is removed <see cref="Enabled"/> should be set to false.
        /// </summary>
        public bool Removed { get; internal set; }
        /// <summary>
        /// The Site id. Guid.
        /// </summary>
        public string SiteId { get; internal set; }
        /// <summary>
        /// The method. 
        /// </summary>
        public eMethod Method { get; internal set; }
        /// <summary>
        /// Content bits.
        /// </summary>
        public eContentBits ContentBits { get; internal set; }

        internal ExcelSensibilityLabel()
        {
            
        }
        internal static ExcelSensibilityLabel CreateFromElement(XmlNamespaceManager nsm, XmlElement element)
        {
            var label = new ExcelSensibilityLabel();
            var helper = XmlHelperFactory.Create(nsm, element);
            label.Id = helper.GetXmlNodeString("@id");
            label.Enabled = helper.GetXmlNodeBool("@enabled");
            label.Removed = helper.GetXmlNodeBool("@removed");
            label.Method = GetMethodEnum(helper.GetXmlNodeString("@method"));
            label.SiteId = helper.GetXmlNodeString("@siteId");
            label.ContentBits = (eContentBits)helper.GetXmlNodeInt("@contentBits", 0);
            return label;
        }

        private static eMethod GetMethodEnum(string method)
        {
            switch(method)
            {
                case "Standard":
                    return eMethod.Standard;
                case "Privileged":
                    return eMethod.Privileged;
                default:
                    return eMethod.Empty;
            }
        }

        internal static IExcelSensibilityLabel CreateFromElement(object nameSpaceManager, XmlElement element)
        {
            throw new NotImplementedException();
        }
    }
}
#endif