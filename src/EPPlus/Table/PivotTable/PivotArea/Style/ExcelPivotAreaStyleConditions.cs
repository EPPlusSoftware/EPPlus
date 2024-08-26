/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/18/2021         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Conditions for a pivot table area style.
    /// </summary>
    public class ExcelPivotAreaStyleConditions
    {
        internal ExcelPivotAreaStyleConditions(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt)
        {
            Fields = new ExcelPivotAreaReferenceCollection(nsm, topNode, pt);
            var xh = XmlHelperFactory.Create(nsm, topNode);
            foreach (XmlElement n in xh.GetNodes("d:references/d:reference"))
            {
                if (n.GetAttribute("field") == "4294967294")
                {
                    DataFields = new ExcelPivotAreaDataFieldReference(nsm, n, pt, -2);
                }
                else
                {
                    Fields.Add(new ExcelPivotAreaReference(nsm, n, pt));
                }
            }

            if(DataFields==null)
            {
                DataFields = new ExcelPivotAreaDataFieldReference(nsm, topNode, pt, -2);
            }
        }
        /// <summary>
        /// Row and column fields that the conditions will apply to. 
        /// </summary>
        public ExcelPivotAreaReferenceCollection Fields 
        { 
            get;  
        }
        /// <summary>
        /// The data field that the conditions will apply to. 
        /// </summary>
        public ExcelPivotAreaDataFieldReference DataFields
        {
            get;
        }
        /// <summary>
        /// Updates the xml. Returns false if all conditions are deleted and the items should be removed.
        /// </summary>
        /// <returns>Returns false if the items should be deleted.</returns>
        internal bool UpdateXml()
        {
            if(DataFields.UpdateXml()==false)
            {
                return false;
            }
            foreach (ExcelPivotAreaReference r in Fields)
            {
                if(r.UpdateXml()==false)
                {
                    return false;
                }
            }
            return true;
        }
    }
}
