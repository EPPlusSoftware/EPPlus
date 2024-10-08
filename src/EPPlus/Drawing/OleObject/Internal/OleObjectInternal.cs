using System.Xml;

namespace OfficeOpenXml.Drawing.OleObject
{

    /*
        <xsd:complexType name="CT_OleObject">
        3072 <xsd:sequence>
        3073 <xsd:element name="objectPr" type="CT_ObjectPr" minOccurs="0" maxOccurs="1"/>
        3074 </xsd:sequence>
        3075 <xsd:attribute name="progId" type="xsd:string" use="optional"/>
        3076 <xsd:attribute name="dvAspect" type="ST_DvAspect" use="optional" default="DVASPECT_CONTENT"/>
        3077 <xsd:attribute name="link" type="s:ST_Xstring" use="optional"/>
        3078 <xsd:attribute name="oleUpdate" type="ST_OleUpdate" use="optional"/>
        3079 <xsd:attribute name="autoLoad" type="xsd:boolean" use="optional" default="false"/>
        3080 <xsd:attribute name="shapeId" type="xsd:unsignedInt" use="required"/>
        3081 <xsd:attribute ref="r:id" use="optional"/>
        3082 </xsd:complexType>
        3083 <xsd:complexType name="CT_ObjectPr">
        3084 <xsd:sequence>
        3085 <xsd:element name="anchor" type="CT_ObjectAnchor" minOccurs="1" maxOccurs="1"/>
        3086 </xsd:sequence>
        3087 <xsd:attribute name="locked" type="xsd:boolean" use="optional" default="true"/>
        3088 <xsd:attribute name="defaultSize" type="xsd:boolean" use="optional" default="true"/>
        3089 <xsd:attribute name="print" type="xsd:boolean" use="optional" default="true"/>
        3090 <xsd:attribute name="disabled" type="xsd:boolean" use="optional" default="false"/>
        3091 <xsd:attribute name="uiObject" type="xsd:boolean" use="optional" default="false"/>
        3092 <xsd:attribute name="autoFill" type="xsd:boolean" use="optional" default="true"/>
        3093 <xsd:attribute name="autoLine" type="xsd:boolean" use="optional" default="true"/>
        3094 <xsd:attribute name="autoPict" type="xsd:boolean" use="optional" default="true"/>
        3095 <xsd:attribute name="macro" type="ST_Formula" use="optional"/>
        3096 <xsd:attribute name="altText" type="s:ST_Xstring" use="optional"/>
        3097 <xsd:attribute name="dde" type="xsd:boolean" use="optional" default="false"/>
        3098 <xsd:attribute ref="r:id" use="optional"/>
    */
    internal class OleObjectInternal : XmlHelper
    {
        internal OleObjectInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }

        internal bool DefaultSize
        {
            get
            {
                return GetXmlNodeBool("d:objectPr/@defaultSize");
            }
            set
            {
                SetXmlNodeBool("d:objectPr/@defaultSize", value);
            }
        }

        ExcelPosition _from = null;
        public ExcelPosition From
        {
            get
            {
                if (_from == null)
                {
                    _from = new ExcelPosition(NameSpaceManager, GetNode("d:objectPr/d:anchor/d:from"), null);
                }
                return _from;
            }
        }
        ExcelPosition _to = null;
        public ExcelPosition To
        {
            get
            {
                if (_to == null)
                {
                    _to = new ExcelPosition(NameSpaceManager, GetNode("d:objectPr/d:anchor/d:to"), null);
                }
                return _to;
            }
        }

        internal bool AutoLoad
        {
            get
            {
                return GetXmlNodeBool("@autoLoad");
            }
            set
            {
                SetXmlNodeBool("@autoLoad", value);
            }
        }

        internal string OleUpdate
        {
            get
            {
                return GetXmlNodeString("@oleUpdate");
            }
            set
            {
                SetXmlNodeString("@oleUpdate", value);
            }
        }

        internal string DvAspect
        {
            get
            {
                return GetXmlNodeString("@dvAspect");
            }
            set
            {
                SetXmlNodeString("@dvAspect", value);
            }
        }

        internal string Link
        {
            get
            {
                return GetXmlNodeString("@link");
            }
            set
            {
                SetXmlNodeString("@link", value);
            }
        }

        internal string ProgId
        {
            get
            {
                return GetXmlNodeString("@progId");
            }
            set
            {
                SetXmlNodeString("@progId", value);
            }
        }

        internal int ShapeId
        {
            get
            {
                return GetXmlNodeInt("@shapeId");
            }
            set
            {
                SetXmlNodeInt("@shapeId", value);
            }
        }

        public string RelationshipId
        {
            get
            {
                return GetXmlNodeString("@r:id");
            }
            set
            {
                SetXmlNodeString("@r:id", value);
            }
        }

        internal void DeleteMe()
        {
            var node = TopNode.ParentNode?.ParentNode;
            if (node?.LocalName == "AlternateContent")
            {
                var parent = node.ParentNode;
                node.ParentNode.RemoveChild(node);
                if(!parent.HasChildNodes)
                {
                    parent.ParentNode.RemoveChild(parent);
                }
            }
            else
            {
                var parent = TopNode.ParentNode;
                TopNode.ParentNode.RemoveChild(TopNode);
                if (!parent.HasChildNodes)
                {
                    parent.ParentNode.RemoveChild(parent);
                }
            }
        }
    }
}
