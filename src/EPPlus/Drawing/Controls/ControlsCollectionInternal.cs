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
using OfficeOpenXml.Drawing.Controls;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
/*
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<mc:Choice Requires="x14">
<controls>
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<mc:Choice Requires="x14">
<control r:id="rId4" name="Button 1" shapeId="1025">
<controlPr macro="[0]!Button1_Click" autoPict="0" autoFill="0" print="0" defaultSize="0">
<anchor sizeWithCells="1" moveWithCells="1">
<from>
<xdr:col>2</xdr:col>
<xdr:colOff>476250</xdr:colOff>
<xdr:row>5</xdr:row>
<xdr:rowOff>47625</xdr:rowOff>
</from>
<to>
<xdr:col>3</xdr:col>
<xdr:colOff>504825</xdr:colOff>
<xdr:row>6</xdr:row>
<xdr:rowOff>47625</xdr:rowOff>
</to>
</anchor>
</controlPr>
</control>
</mc:Choice>
</mc:AlternateContent>
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
</controls>
</mc:Choice>
</mc:AlternateContent> 
 */
namespace OfficeOpenXml
{
    internal class ControlsCollectionInternal : XmlHelper, IEnumerable<ControlInternal>
    {
        internal List<ControlInternal> _list=new List<ControlInternal>();
        internal ControlsCollectionInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            var nodes = GetNodes("mc:AlternateContent/mc:Choice/controls/mc:AlternateContent/d:choice/controls");
            foreach(XmlNode node in nodes)
            {
                _list.Add(new ControlInternal(NameSpaceManager, node));
            }
        }

        public IEnumerator<ControlInternal> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}