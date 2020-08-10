/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/29/2020         EPPlus Software AB       Threaded comments
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// A person in the context of ThreadedComments.
    /// Might refer to an <see cref="IdentityProvider"/>, see property ProviderId.
    /// </summary>
    public class ExcelThreadedCommentPerson : XmlHelper, IEqualityComparer<ExcelThreadedCommentPerson>
    {
        internal static string NewId()
        {
            var guid = Guid.NewGuid();
            return "{" + guid.ToString().ToUpper() + "}";
        }

        internal ExcelThreadedCommentPerson(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            TopNode = topNode;
            SchemaNodeOrder = new string[] { "displayName", "id", "userId", "providerId" };
        }

        /// <summary>
        /// Unique Id of the person
        /// </summary>
        public string Id
        {
            get { return GetXmlNodeString("@id"); }
            set { SetXmlNodeString("@id", value); }
        }

        /// <summary>
        /// Display name of the person
        /// </summary>
        public string DisplayName
        {
            get { return GetXmlNodeString("@displayName"); }
            set { SetXmlNodeString("@displayName", value); }
        }

        /// <summary>
        /// See the documentation of the members of the <see cref="IdentityProvider"/> enum and
        /// Microsofts documentation at https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968
        /// </summary>
        public string UserId
        {
            get { return GetXmlNodeString("@userId"); }
            set { SetXmlNodeString("@userId", value); }
        }

        /// <summary>
        /// See the documentation of the members of the <see cref="IdentityProvider"/> enum and
        /// Microsofts documentation at https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968
        /// </summary>
        public IdentityProvider ProviderId
        {
            get 
            { 
                var id = GetXmlNodeString("@providerId");
                if (string.IsNullOrEmpty(UserId) && UserId == "AD") throw new InvalidOperationException("Cannot get ProviderId when UserId is not set");
                switch(id)
                {
                    case "Windows Live":
                        return IdentityProvider.WindowsLiveId;
                    case "PeoplePicker":
                        return IdentityProvider.PeoplePicker;
                    case "AD":
                        if (UserId.Contains("::"))
                            return IdentityProvider.Office365;
                        return IdentityProvider.ActiveDirectory;
                    default:
                        return IdentityProvider.NoProvider;
                }
            
            }
            set 
            {
                switch(value)
                {
                    case IdentityProvider.ActiveDirectory:
                        SetXmlNodeString("@providerId", "AD");
                        break;
                    case IdentityProvider.WindowsLiveId:
                        SetXmlNodeString("@providerId", "Windows Live");
                        break;
                    case IdentityProvider.Office365:
                        SetXmlNodeString("@providerId", "AD");
                        break;
                    case IdentityProvider.PeoplePicker:
                        SetXmlNodeString("@providerId", "PeoplePicker");
                        break;
                    default:
                        SetXmlNodeString("@providerId", "None");
                        break;
                }
            }
        }

        public bool Equals(ExcelThreadedCommentPerson x, ExcelThreadedCommentPerson y)
        {
            if (x == null && y == null) return true;
            if (x == null ^ y == null) return false;
            if (x.UserId == y.UserId) return true;
            return false;
        }

        public int GetHashCode(ExcelThreadedCommentPerson obj)
        {
            return obj.GetHashCode();
        }

        public override string ToString()
        {
            return DisplayName + " (id: " + UserId + ")";
        }
    }
}
