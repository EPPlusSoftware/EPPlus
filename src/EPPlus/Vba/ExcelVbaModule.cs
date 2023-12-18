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
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.VBA
{
    internal delegate void ModuleNameChange(string value);

    /// <summary>
    /// A VBA code module. 
    /// </summary>
    public class ExcelVBAModule
    {
        string _name = "";
        ModuleNameChange _nameChangeCallback = null;
        private static readonly char[] _nonValidChars = new char[] { '!', '\\', '"', '@', '#', '$', '%', '&', '/', '{', '}', '[', ']', '(', ')', '<', '>', '=', '+', '-', '?', '`', '~', '^', '\'', '*', ';', ':', ' ', '.', ' ', '«', '»' };
        internal ExcelVBAModule()
        {
            Attributes = new ExcelVbaModuleAttributesCollection();
        }
        internal ExcelVBAModule(ModuleNameChange nameChangeCallback) :
            this()
        {
            _nameChangeCallback = nameChangeCallback;
        }
        /// <summary>
        /// The name of the module
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                if (!IsValidModuleName(value))
                {
                    throw (new InvalidOperationException("Name contains invalid characters"));
                }
                if (value != _name)
                {
                    _name = value;
                    streamName = value;
                    if (_nameChangeCallback != null)
                    {
                        _nameChangeCallback(value);
                    }
                }
            }
        }

        /// <summary>
        /// Module name unicode
        /// </summary>
        internal string NameUnicode { get; set; }

        internal static bool IsValidModuleName(string name)
        {
            if (string.IsNullOrEmpty(name) ||   //Not null or empty
               char.IsLetter(name[0]) == false ||        //Don't start with a number or underscore
               name.Any(x => x < 0x30 || IsAbove255AndNotLetter(x) || _nonValidChars.Contains(x))) //Don't contain invalid chars. Allow unicode
            {
                return false;
            }
            return true;
        }

        static bool IsAbove255AndNotLetter(char c)
        {
            if (c > 255)
            {
                return (char.IsLetter(c) == false);
            }
            return false;
        }
        /// <summary>
        /// A description of the module
        /// </summary>
        public string Description { get; set; }
        private string _code = "";
        /// <summary>
        /// The code without any module level attributes.
        /// <remarks>Can contain function level attributes.</remarks> 
        /// </summary>
        public string Code {
            get
            {
                return _code;
            }
            set
            {
                if (value.StartsWith("Attribute", StringComparison.OrdinalIgnoreCase) || value.StartsWith("VERSION", StringComparison.OrdinalIgnoreCase))
                {
                    throw (new InvalidOperationException("Code can't start with an Attribute or VERSION keyword. Attributes can be accessed through the Attributes collection."));
                }
                _code = value;
            }
        }
        /// <summary>
        /// A reference to the helpfile
        /// </summary>
        public int HelpContext { get; set; }
        /// <summary>
        /// Module level attributes.
        /// </summary>
        public ExcelVbaModuleAttributesCollection Attributes { get; internal set; }
        /// <summary>
        /// Type of module
        /// </summary>
        public eModuleType Type { get; internal set; }
        /// <summary>
        /// If the module is readonly
        /// </summary>
        public bool ReadOnly { get; set; }
        /// <summary>
        /// If the module is private
        /// </summary>
        public bool Private { get; set; }
        internal string streamName { get; set; }
        internal ushort Cookie { get; set; }
        internal uint ModuleOffset { get; set; }
        internal string ClassID { get; set; }
        /// <summary>
        /// Converts the object to a string
        /// </summary>
        /// <returns>The name of the VBA module</returns>
        public override string ToString()
        {
            return Name;
        }
    }
}
