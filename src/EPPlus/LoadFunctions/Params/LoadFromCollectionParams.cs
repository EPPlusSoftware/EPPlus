/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters for the LoadFromCollection method
    /// </summary>
    public class LoadFromCollectionParams : LoadFunctionFunctionParamsBase
    {
        private readonly Dictionary<string, IEnumerable<string>> _dictionaryKeys = new Dictionary<string, IEnumerable<string>>();
        private readonly string DefaultDictionaryKeyId = Guid.NewGuid().ToString("N");
        /// <summary>
        /// Default value for the BindingFlags property
        /// </summary>
        public const BindingFlags DefaultBindingFlags = BindingFlags.Public | BindingFlags.Instance;

        /// <summary>
        /// The <see cref="BindingFlags"/> used when reading properties via reflection.
        /// </summary>
        public BindingFlags BindingFlags { get; set; } = DefaultBindingFlags;

        /// <summary>
        /// If not null, this specifies the members that should be used. Any member not present will be ignored.
        /// </summary>
        public MemberInfo[] Members { get; set; }

        /// <summary>
        /// Sets how headers should be parsed before added to the worksheet, see <see cref="HeaderParsingTypes"/>
        /// </summary>
        public HeaderParsingTypes HeaderParsingType { get; set; } = HeaderParsingTypes.UnderscoreToSpace;

        /// <summary>
        /// Register keys to a property decorated with the <see cref="EPPlusDictionaryColumnAttribute"/>. These will also
        /// be used to create the column for this property.
        /// The <paramref name="keyId"/> should map to the <see cref="EPPlusDictionaryColumnAttribute.KeyId">KeyId property of the attribute.</see>
        /// </summary>
        /// <param name="keyId">Key id used to store this set of keys</param>
        /// <param name="keys">Keys for the </param>
        public void RegisterDictionaryKeys(string keyId, IEnumerable<string> keys)
        {
            if(string.IsNullOrEmpty(keyId))
            {
                throw new ArgumentNullException($"{nameof(keyId)} cannot be null or empty");
            }
            if(_dictionaryKeys.ContainsKey(keyId))
            {
                throw new InvalidOperationException($"The keyId '{keyId}' has already been used.");
            }
            if(keys == null)
            {
                throw new ArgumentNullException(nameof(keys));
            }
            if(!keys.Any())
            {
                throw new ArgumentException($"Parameter {nameof(keys)} cannot be empty");
            }
            _dictionaryKeys.Add(keyId, keys);
        }

        /// <summary>
        /// Registers default keys for properties decorated with the <see cref="EPPlusDictionaryColumnAttribute"/>. These will also
        /// be used to create the column for this property.
        /// </summary>
        /// <param name="keys">The keys to register</param>
        public void RegisterDictionaryKeys(IEnumerable<string> keys)
        {
            RegisterDictionaryKeys(DefaultDictionaryKeyId, keys);
        }

        internal IEnumerable<string> GetDictionaryKeys(string keyId)
        {
            if(_dictionaryKeys.ContainsKey(keyId)) return _dictionaryKeys[keyId];
            return Enumerable.Empty<string>();
        }

        internal IEnumerable<string> GetDefaultDictionaryKeys()
        {
            return GetDictionaryKeys(DefaultDictionaryKeyId);
        }
    }
}
