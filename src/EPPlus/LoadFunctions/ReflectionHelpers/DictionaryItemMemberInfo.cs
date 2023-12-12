﻿/*************************************************************************************************
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
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal class DictionaryItemMemberInfo : MemberInfo
    {
        public DictionaryItemMemberInfo(string key)
        {
            _key = key;
        }

        private readonly string _key;
        public override Type DeclaringType => typeof(Dictionary<string, string>);

        public override MemberTypes MemberType => MemberTypes.Custom;

        public override string Name => _key;

        public override Type ReflectedType => typeof(string);

        public override object[] GetCustomAttributes(bool inherit)
        {
            return Enumerable.Empty<object>().ToArray();
        }

        public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        {
            return Enumerable.Empty<object>().ToArray();
        }

        public override bool IsDefined(Type attributeType, bool inherit)
        {
            return false;
        }
    }
}
