/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/29/2023         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Represents a cell protection properties used for differential style formatting.
    /// </summary>
    public class ExcelDxfProtection : DxfStyleBase
    {
        internal ExcelDxfProtection(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
        }
        /// <summary>
        /// If the cell is locked when the worksheet is protected. 
        /// </summary>
        public bool? Locked { get; set; }
        /// <summary>
        /// If the cells formulas are hidden when the worksheet is protected. 
        /// </summary>
        public bool? Hidden { get; set; }
        /// <summary>
        /// If the dxf style has any values set.
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return Locked.HasValue || 
                       Hidden.HasValue;
            }
        }

        internal override string Id 
        {
            get
            {
                return GetAsString(Hidden) + "|" +
                       GetAsString(Locked);
            }
        }

        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {

            Locked = null;
            Hidden = null;
        }

        internal override DxfStyleBase Clone()
        {
            return new ExcelDxfProtection(_styles, _callback) 
            { 
                Locked = Locked,
                Hidden = Hidden
            };
        }

        internal override void CreateNodes(XmlHelper helper, string path)
        {
            SetValueBool(helper, path + "/@locked", Locked);
            SetValueBool(helper, path + "/@hidden", Hidden);
        }

        internal override void SetStyle()
        {
            if(_callback!=null)
            {
                _callback.Invoke(eStyleClass.Style, eStyleProperty.Locked, Locked);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.Hidden, Hidden);
            }
        }
    }
}
