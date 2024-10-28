/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing.OleObject;

namespace OfficeOpenXml
{
    internal class OleObjectsCollectionInternal : XmlHelper, IEnumerable<KeyValuePair<int, OleObjectInternal>>
    {
        internal Dictionary<int ,OleObjectInternal> _dict = new Dictionary<int, OleObjectInternal>();

        internal OleObjectsCollectionInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            var nodes = GetNodes("d:oleObjects/mc:AlternateContent/mc:Choice/d:oleObject");
            foreach (XmlNode node in nodes)
            {
                int shapeId = int.Parse(node.Attributes["shapeId"].Value);
                _dict.Add(shapeId, new OleObjectInternal(NameSpaceManager, node));
            }
        }

        internal OleObjectInternal GetOleObjectByShapeId(int shapeId)
        {
            return _dict.FirstOrDefault(x => x.Key == shapeId).Value;
        }

        public IEnumerator<KeyValuePair<int, OleObjectInternal>> GetEnumerator()
        {
            return _dict.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _dict.GetEnumerator();
        }
    }
}
