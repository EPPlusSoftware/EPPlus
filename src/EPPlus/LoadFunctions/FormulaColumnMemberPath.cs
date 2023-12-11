/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/7/2023         EPPlus Software AB       EPPlus 7.0.4
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Attributes;

namespace OfficeOpenXml.LoadFunctions
{
    internal class FormulaColumnMemberPath : MemberPathBase
    {
        public FormulaColumnMemberPath(EpplusFormulaTableColumnAttribute attr)
        {
            _attr = attr;
            Init();
        }

        private readonly EpplusFormulaTableColumnAttribute _attr;

        private void Init()
        {
            var item = new MemberPathItem(_attr);
            _members.Add(item);
        }

        public override bool IsFormulaColumn => true;
        public override string GetHeader()
        {
            return _attr.Header;
        }

        internal override string GetPath()
        {
            return string.Empty;
        }
    }
}
