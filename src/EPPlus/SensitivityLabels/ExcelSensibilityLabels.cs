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
using OfficeOpenXml.Interfaces.SensitivityLabels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeOpenXml.SensitivityLabels
{
    public class ExcelSensibilityLabels
    {
        static ISensitivityLabelHandler _sensibilityLabelHandler=null;
        /// <summary>
        /// If you want your workbooks to be marked with sensibility lables, you can add a handler for authentication, encryption and decryption using the Microsoft Information Protection SDK.
        /// For more information
        /// </summary>
        public static ISensitivityLabelHandler SensibilityLabelHandler 
        {
            get
            {
                return _sensibilityLabelHandler;
            }
            set
            {
                if (value != null && value != _sensibilityLabelHandler)
                {
                    value.InitAsync();
                }
                _sensibilityLabelHandler = value;
            } 
        }

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
            foreach(XmlElement element in xml.SelectNodes("/clbl:labelList/clbl:label", _nsm))
            {
                Labels.Add(ExcelSensibilityLabel.CreateFromElement(_nsm, element));
            }

            SensibilityLabelHandler.UpdateLabelList(Labels, _pck.Id);
        }

        /// <summary>
        /// Changes the active sensibility label. This will overwrite the active sensibility label and any settings derived from it.
        /// </summary>
        /// <param name="id">The id for the sensibility label without heading and traling brackets.</param>
        public void SetActiveLabelById(string id)
        {
            var lbls = SensibilityLabelHandler.GetLabels();

            var lbl = lbls.FirstOrDefault(x => x.Id.Equals(id, StringComparison.CurrentCultureIgnoreCase));
            if (lbl == null)
            {
                throw (new ArgumentException($"Sensitivity label with id:{id} does not exist."));
            }
            AddLabel(lbl);
        }
        /// <summary>
        /// Changes the active sensibility label. This will overwrite the active sensibility label and any settings derived from it.
        /// </summary>
        /// <param name="name"></param>
        public void SetActiveLabelByName(string name)
        {
            if (SensibilityLabelHandler == null)
            {
                throw (new MissingSensibilityHandlerException("No sensibility label handler is set. Please set the property ExcelSensibilityLabels.SensibilityLabelHandler"));
            }

            var lbls = SensibilityLabelHandler.GetLabels();

            var lbl = lbls.FirstOrDefault(x => x.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase));
            if (lbl == null)
            {
                throw (new ArgumentException($"Sensitivity label {name} does not exist."));
            }
            AddLabel(lbl);
        }

        private void AddLabel(IExcelSensibilityLabel lbl)
        {
            var existingLabel = Labels.FirstOrDefault(x => x.Id.Equals(lbl.Id, StringComparison.InvariantCultureIgnoreCase));
            if (existingLabel == null || existingLabel.Enabled == false)
            {
                ActiveLabelId = lbl.Id;
                ProtectionInformation = null;
                _pck.Encryption.Version = EncryptionVersion.ProtectedBySensibilityLabel;

                Labels.Add(new ExcelSensibilityLabel
                {
                    Id = lbl.Id,
                    Name = lbl.Name,
                    Description = lbl.Description,
                    Color = lbl.Color,
                    SiteId = lbl.SiteId,
                    ContentBits = lbl.ContentBits,
                    Enabled = true,
                    Removed = false,
                    Method = eMethod.Privileged
                });
            }
        }

        /// <summary>
        /// Contains sensitivity labels for the package. 
        /// Only the last sensitivity will be applied on save.
        /// </summary>
        public EPPlusReadOnlyList<IExcelSensibilityLabel> Labels
        {
            get;
        }
        /// <summary>
        /// Property used by the <see cref="SensibilityLabelHandler"/> to store information about the sensibility label. 
        /// The information is passed when calling the DecryptPackageAsync and ApplyLabelAndSavePackageAsync methods.
        /// <seealso cref="ISensitivityLabelHandler.DecryptPackageAsync(System.IO.MemoryStream, string)"/>
        /// <seealso cref="ISensitivityLabelHandler.ApplyLabelAndSavePackageAsync(IDecryptedPackage, string)"/>
        /// </summary>
        public object ProtectionInformation { get; set; }
        /// <summary>
        /// The latest set sensitivity label.
        /// </summary>
        internal string ActiveLabelId { get; set; }

#if (!NET35)
        internal async Task<MemoryStream> ApplyLabel(byte[] bytes)
        {
            var decryptionInfo = new EPPlusDecryptionInfo()
            {
                PackageStream = new MemoryStream(bytes),
                ProtectionInformation = _pck.SensibilityLabels.ProtectionInformation,
                ActiveLabelId = _pck.SensibilityLabels.ActiveLabelId
            };
            return await ExcelSensibilityLabels.SensibilityLabelHandler.ApplyLabelAndSavePackageAsync(decryptionInfo, _pck.Id);
        }
#endif
    }

    [DebuggerDisplay("Name: {Name}")]
    public class ExcelSensibilityLabel : IExcelSensibilityLabel, IExcelSensibilityLabelUpdate
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

        /// <summary>
        /// The color of the label.
        /// </summary>
        public string Color { get; internal set; }

        public string Tooltip { get; internal set; }

        /// <summary>
        /// The parent label, if any.
        /// </summary>
        public IExcelSensibilityLabel Parent { get; internal set; }
        /// <summary>
        /// Update properties from the handler
        /// </summary>
        /// <param name="name">The name of the label</param>
        /// <param name="tooltip">The tooltip for the label</param>
        /// <param name="description">The desription</param>
        /// <param name="color">The RGB color in hex</param>
        /// <param name="parent">The id of the parent of the label.</param>
        public void Update(string name, string tooltip, string description, string color, IExcelSensibilityLabel parent)
        {
            Name = name;
            Tooltip = tooltip;
            Description = description;
            Color = color;
            Parent = parent;
        }
        internal static ExcelSensibilityLabel CreateFromElement(XmlNamespaceManager nsm, XmlElement element)
        {
            var label = new ExcelSensibilityLabel();
            var helper = XmlHelperFactory.Create(nsm, element);
            label.Id = helper.GetXmlNodeString("@id").TrimStart('{').TrimEnd('}'); // Remove the brackets, so it matches the id in the MIPS api.
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
    }
}
#endif