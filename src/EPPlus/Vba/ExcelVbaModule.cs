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
        private static readonly char[] _nonValidChars = new char[] { '!', '\\', '"', '@', '#', '$', '%', '&', '/', '{', '}', '[', ']', '(', ')', '<', '>', '=', '+', '-', '?', '`', '~', '^', '\'', '*', ';', ':' };
        //private const string _validModulePattern = "^[a-zA-Z][a-zA-Z0-9_ ]*$";
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
                if (value.Any(c => c > 255))
                {
                    throw (new InvalidOperationException("Vba module names can't contain unicode characters"));
                }
                if(!IsValidModuleName(value))
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
        internal static bool IsValidModuleName(string name)
        {
            //return Regex.IsMatch(name, _validModulePattern);
            if (string.IsNullOrEmpty(name) ||           //Not null or empty
               (name[0]>='0' && name[0]<=9) ||          //Don't start with a number
               name.Any(x=>x<0x20  || x > 255 || _nonValidChars.Contains(x)))      //Don't contain invalid or unicode chars 
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// A description of the module
        /// </summary>
        public string Description { get; set; }
        private string _code="";
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
                if(value.StartsWith("Attribute",StringComparison.OrdinalIgnoreCase) || value.StartsWith("VERSION",StringComparison.OrdinalIgnoreCase))
                {
                    throw(new InvalidOperationException("Code can't start with an Attribute or VERSION keyword. Attributes can be accessed through the Attributes collection."));
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
