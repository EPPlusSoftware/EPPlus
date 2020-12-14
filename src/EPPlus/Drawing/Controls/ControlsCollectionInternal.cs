/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/01/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Controls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
namespace OfficeOpenXml
{
    internal class ControlsCollectionInternal : XmlHelper, IEnumerable<ControlInternal>
    {
        private List<ControlInternal> _list=new List<ControlInternal>();
        internal ControlsCollectionInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            var nodes = GetNodes("mc:AlternateContent/mc:Choice/d:controls/mc:AlternateContent/mc:Choice/d:control");
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

        internal ControlInternal GetControlByShapeId(int shapeId)
        {
            return _list.FirstOrDefault(x => x.ShapeId == shapeId);
        }
    }
}