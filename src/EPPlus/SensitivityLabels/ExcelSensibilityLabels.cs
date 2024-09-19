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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
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

            ExcelPackage.SensibilityLabelHandler.UpdateLabelList(Labels, _pck.Id);
        }
        public void SetActiveLabelById(string id)
        {
            SetActiveLabelById(id, null);
        }
        /// <summary>
        /// Changes the active sensibility label. This will overwrite the active sensibility label and any settings derived from it, if the <see cref="ExcelPackage.SensibilityLabelHandler"/> is used.
        /// </summary>
        /// <param name="id">The id for the sensibility label without heading and traling brackets.</param>
        /// <param name="siteId">The site id for the tenent. This id is used if no <see cref="ExcelPackage.SensibilityLabelHandler" /> is set.</param>
        /// <param name="method">Sets the method property. This property is used if no <see cref="ExcelPackage.SensibilityLabelHandler" /> is set.</param>
        /// <param name="clearLabels">If true, will not clear the <see cref="Labels"/> collection. If false, will set the active sensibility label(if any) to Readonly=true and Enabled=false</param>
        public void SetActiveLabelById(string id, string siteId, eMethod method=eMethod.Privileged, bool clearLabels=true)
        {
            if(string.IsNullOrEmpty(id))
            {
                throw new ArgumentNullException("id", $"Sensibility Label Id can not be null or empty.");
            }

            if (ExcelPackage.SensibilityLabelHandler == null)
            {
                if (string.IsNullOrEmpty(siteId))
                {
                    throw new ArgumentNullException("id", $"Sensibility Label Id can not be null or empty.");
                }
                AddLabel(new ExcelSensibilityLabel() { Id= id,  SiteId = siteId, Method=method, Enabled=true, Removed=false}, clearLabels);
            }
            else
            {
                var lbls = ExcelPackage.SensibilityLabelHandler.GetLabels(_pck.Id);

                var lbl = lbls.FirstOrDefault(x => x.Id.Equals(id, StringComparison.CurrentCultureIgnoreCase));
                if (lbl == null)
                {
                    throw new ArgumentException($"Sensitivity label with id:{id} does not exist.");
                }
                AddLabel(lbl, clearLabels);
            }
        }
        /// <summary>
        /// Changes the active sensibility label. This will overwrite the active sensibility label and any settings derived from it.
        /// </summary>
        /// <param name="name"></param>
        public void SetActiveLabelByName(string name)
        {
            if (ExcelPackage.SensibilityLabelHandler == null)
            {
                throw (new MissingSensibilityHandlerException("No sensibility label handler is set. Please set the property ExcelPackage.SensibilityLabelHandler"));
            }

            var lbls = ExcelPackage.SensibilityLabelHandler.GetLabels(_pck.Id);

            var lbl = lbls.FirstOrDefault(x => x.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase));
            if (lbl == null)
            {
                throw (new ArgumentException($"Sensitivity label with name: {name} does not exist."));
            }
            AddLabel(lbl, false);
        }

        private void AddLabel(IExcelSensibilityLabel lbl, bool clearLabels)
        {
            var existingLabel = Labels.FirstOrDefault(x => x.Id.Equals(lbl.Id, StringComparison.InvariantCultureIgnoreCase));
            if (existingLabel == null || existingLabel.Enabled == false)
            {
                ActiveLabelId = lbl.Id;
                ProtectionInformation = null;
                _pck.Encryption.Version = EncryptionVersion.ProtectedBySensibilityLabel;
                if (clearLabels)
                {
                    Labels._list.Clear();
                }
                else
                {
                    Labels._list.ForEach(x => { var l = (ExcelSensibilityLabel)x; l.Enabled = false; l.Removed = true; });
                }
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
                    Method = lbl.Method
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
        /// Property used by the <see cref="ExcelPackage.SensibilityLabelHandler"/> to store information about the sensibility label. 
        /// The information is passed when calling the DecryptPackageAsync and ApplyLabelAndSavePackageAsync methods.
        /// <seealso cref="ISensitivityLabelHandler.DecryptPackageAsync(MemoryStream, string)"/>
        /// <seealso cref="ISensitivityLabelHandler.ApplyLabelAndSavePackageAsync(IPackageInfo, string)"/>
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
            return await ExcelPackage.SensibilityLabelHandler.ApplyLabelAndSavePackageAsync(decryptionInfo, _pck.Id);
        }

        internal void SaveToXml()
        {
            if (Labels.Count == 0)
            {
                return;
            }
            var part = _pck.ZipPackage.GetPartByContentType(ContentTypes.contentTypeClassificationLabels);
            var xml = new XmlDocument();
            if (part == null)
            {
                var uri = new Uri("/docMetadata/LabelInfo.xml", UriKind.Relative);
                part = _pck.ZipPackage.CreatePart(uri, ContentTypes.contentTypeClassificationLabels, CompressionLevel.Default);
                _pck.ZipPackage.CreateRelationship(uri.OriginalString, Packaging.TargetMode.Internal, "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels");                
            }
            xml= new XmlDocument();
            xml.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><clbl:labelList xmlns:clbl=\"http://schemas.microsoft.com/office/2020/mipLabelMetadata\" />");
            var topNode = xml.DocumentElement;
            foreach (var label in Labels)
            {
                var labelElement = xml.CreateElement("clbl", "label", "http://schemas.microsoft.com/office/2020/mipLabelMetadata");                
                labelElement.SetAttribute("id", FormatGuid(label.Id));
                labelElement.SetAttribute("removed", label.Removed ? "1" : "0");
                labelElement.SetAttribute("enabled", label.Enabled ? "1" : "0");
                labelElement.SetAttribute("siteId", FormatGuid(label.SiteId));
                if (label.Method != eMethod.Empty)
                {
                    labelElement.SetAttribute("method", label.Method.ToString());
                }
                labelElement.SetAttribute("contentBits", ((int)label.ContentBits).ToString(CultureInfo.InvariantCulture));
                topNode.AppendChild(labelElement);
            }
            var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(stream);
            stream.Flush();
        }

        private string FormatGuid(string guid)
        {
            if(guid == null || guid.Length!=36) return guid;
            if (guid[0] != '{')
            {
                guid = "{" + guid;
            }
            if (guid.Last() != '}')
            {
                guid += "}";
            }
            return guid;
        }
#endif
    }

    /// <summary>
    /// Represents a sensitivity label that can be applied to a package.
    /// </summary>
    [DebuggerDisplay("Name: {Name}")]   
    public class ExcelSensibilityLabel : IExcelSensibilityLabel, IExcelSensibilityLabelUpdate
    {
        /// <summary>
        /// The sensitivity label id. Guid.
        /// </summary>
        public string Id { get; internal set; }
        /// <summary>
        /// The name of the sensibility label. If no <see cref="ExcelPackage.SensibilityLabelHandler"/> is set this property will always be empty.
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// The description of the sensibility label. If no <see cref="ExcelPackage.SensibilityLabelHandler"/> is set this property will always be empty.
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
        /// The Azure AD site id. Guid.
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

        /// <summary>
        /// The description of the sensibility label for the end user.
        /// </summary>
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