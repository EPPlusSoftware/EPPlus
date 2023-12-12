/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2023         EPPlus Software AB           EPPlus 7.0.2
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class UsedTypesScanner
    {
        public UsedTypesScanner(Type outerType)
        {
            _outerType = outerType;
        }

        private readonly Type _outerType;

        private void ScanType(HashSet<Type> types, Type typeToScan)
        {
            if(!types.Contains(typeToScan))
            {
                types.Add(typeToScan);
            }
            var properties = typeToScan.GetProperties();
            foreach(var property in properties)
            {
                if(property.HasAttributeOfType<EpplusNestedTableColumnAttribute>())
                {
                    var propType = property.PropertyType;
                    if(!types.Contains(propType))
                    {
                        types.Add(propType);
                        ScanType(types, propType);
                    }
                }
            }
        }

        public void ValidateMembers(IEnumerable<MemberInfo> members)
        {
            if (members == null || !members.Any()) return;
            var typesToValidate = members.Select(m => m.DeclaringType).Distinct();
            ValidateTypes(typesToValidate);
        }

        private void ValidateTypes(IEnumerable<Type> typesToValidate)
        {
            var includedTypes = new HashSet<Type>();
            ScanType(includedTypes, _outerType);
            foreach(var typeToValidate in typesToValidate)
            {
                var isValid = false;
                Type invalidType = null;
                foreach (var includedType in includedTypes)
                {

                    if (typeToValidate == includedType
                        || typeToValidate.IsAssignableFrom(includedType)
                        || typeToValidate.IsSubclassOf(includedType))
                    {
                        isValid = true;
                        break;
                    }
                    else
                    {
                        invalidType = typeToValidate;
                    }
                }

                if (!isValid) throw new InvalidCastException($"Invalid Declaring type in the members array ({invalidType.FullName})");
            }
        }
    }
}
