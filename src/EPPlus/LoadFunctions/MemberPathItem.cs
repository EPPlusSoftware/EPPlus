using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class MemberPathItem
    {
        public MemberPathItem(MemberInfo member)
        {
            Member = member;
        }

        public MemberInfo Member { get; set; }

        public bool IsNestedProperty { get; set; }
    }
}
