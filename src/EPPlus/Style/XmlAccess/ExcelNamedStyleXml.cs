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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for named styles
    /// </summary>
    public sealed class ExcelNamedStyleXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelNamedStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            _styles = styles;
            BuildInId = int.MinValue;
        }
        internal ExcelNamedStyleXml(XmlNamespaceManager NameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(NameSpaceManager, topNode)
        {
            StyleXfId = GetXmlNodeInt(idPath);
            Name = GetXmlNodeString(namePath);
            BuildInId = GetXmlNodeInt(buildInIdPath);
            CustomBuildin = GetXmlNodeBool(customBuiltinPath);

            _styles = styles;
            _style = new ExcelStyle(styles, styles.NamedStylePropertyChange, -1, Name, _styleXfId);
        }
        internal override string Id
        {
            get
            {
                return Name;
            }
        }
        int _styleXfId=0;
        const string idPath = "@xfId";
        /// <summary>
        /// Named style index
        /// </summary>
        public int StyleXfId
        {
            get
            {
                return _styleXfId;
            }
            set
            {
                _styleXfId = value;
            }
        }
        int _xfId = int.MinValue;
        /// <summary>
        /// Style index
        /// </summary>
        internal int XfId
        {
            get
            {
                return _xfId;
            }
            set
            {
                _xfId = value;
            }
        }
        const string buildInIdPath = "@builtinId";
        /// <summary>
        /// The build in Id for the named style
        /// </summary>
        public int BuildInId { get; set; }
        const string customBuiltinPath = "@customBuiltin";
        /// <summary>
        /// Indicates if this built-in cell style has been customized
        /// </summary>
        public bool CustomBuildin { get; set; }
        const string namePath = "@name";
        string _name;
        /// <summary>
        /// Name of the style
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            internal set
            {
                _name = value;
            }
        }
        ExcelStyle _style = null;
        /// <summary>
        /// The style object
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _style;
            }
            internal set
            {
                _style = value;
            }
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString(namePath, _name);
            SetXmlNodeString("@xfId", _styles.CellStyleXfs[StyleXfId].newID.ToString());
            if (BuildInId>=0) SetXmlNodeString("@builtinId", BuildInId.ToString());
            if(CustomBuildin) SetXmlNodeBool(customBuiltinPath, true);
            return TopNode;            
        }
    }
}
