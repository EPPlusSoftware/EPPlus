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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Collection of sparklines
    /// </summary>
    public class ExcelSparklineCollection : IEnumerable<ExcelSparkline>
    {
        ExcelSparklineGroup _slg;
        List<ExcelSparkline> _lst;
        internal ExcelSparklineCollection(ExcelSparklineGroup slg)
        {
            _slg = slg;
            _lst = new List<ExcelSparkline>();
            LoadSparklines();
        }
        const string _topPath = "x14:sparklines/x14:sparkline";
        /// <summary>
        /// Number of sparklines in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _lst.Count;
            }            
        }

        private void LoadSparklines()
        {
            var grps=_slg.TopNode.SelectNodes(_topPath, _slg.NameSpaceManager);
            foreach(XmlElement grp in grps)
            {
                _lst.Add(new ExcelSparkline(_slg.NameSpaceManager, grp));
            }
        }
        /// <summary>
        /// Returns the sparklinegroup at the specified position.  
        /// </summary>
        /// <param name="index">The position of the Sparklinegroup. 0-base</param>
        /// <returns></returns>
        public ExcelSparkline this[int index]
        {
            get
            {
                return (_lst[index]);
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelSparkline> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        internal void Add(ExcelCellAddress cell, string worksheetName, ExcelAddressBase sqref)
        {
            var sparkline = _slg.TopNode.OwnerDocument.CreateElement("x14","sparkline", ExcelPackage.schemaMainX14);            
            var sls = _slg.TopNode.SelectSingleNode("x14:sparklines", _slg.NameSpaceManager);

            sls.AppendChild(sparkline);
            _slg.TopNode.AppendChild(sls);
            var sl = new ExcelSparkline(_slg.NameSpaceManager, sparkline);
            sl.Cell = cell;
            sl.RangeAddress = sqref;
            _lst.Add(sl);
        }
        internal void Remove(ExcelSparkline sparkline)
        {
            sparkline.TopNode.ParentNode.RemoveChild(sparkline.TopNode);
            _lst.Remove(sparkline);
        }
    }
}
