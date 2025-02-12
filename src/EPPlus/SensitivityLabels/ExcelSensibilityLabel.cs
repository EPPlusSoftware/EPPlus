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
using OfficeOpenXml.Interfaces.SensitivityLabels;
using System.Diagnostics;
using System.Xml;

namespace OfficeOpenXml.SensitivityLabels
{
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