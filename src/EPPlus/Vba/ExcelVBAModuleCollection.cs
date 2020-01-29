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

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// Collection class for VBA modules
    /// </summary>
    public class ExcelVbaModuleCollection : ExcelVBACollectionBase<ExcelVBAModule>
    {
        ExcelVbaProject _project;
        internal ExcelVbaModuleCollection (ExcelVbaProject project)
	    {
            _project=project;
	    }
        internal void Add(ExcelVBAModule Item)
        {
            _list.Add(Item);
        }
        /// <summary>
        /// Adds a new VBA Module
        /// </summary>
        /// <param name="Name">The name of the module</param>
        /// <returns>The module object</returns>
        public ExcelVBAModule AddModule(string Name)
        {
            if (this[Name] != null)
            {
                throw(new ArgumentException("Vba modulename already exist."));
            }
            var m = new ExcelVBAModule();
            m.Name = Name;
            m.Type = eModuleType.Module;
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Name", Value = Name, DataType = eAttributeDataType.String });
            m.Type = eModuleType.Module;
            _list.Add(m);
            return m;
        }
        /// <summary>
        /// Adds a new VBA class
        /// </summary>
        /// <param name="Name">The name of the class</param>
        /// <param name="Exposed">Private or Public not createble</param>
        /// <returns>The class object</returns>
        public ExcelVBAModule AddClass(string Name, bool Exposed)
        {
            var m = new ExcelVBAModule();
            m.Name = Name;            
            m.Type = eModuleType.Class;
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Name", Value = Name, DataType = eAttributeDataType.String });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Base", Value = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}", DataType = eAttributeDataType.String });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_GlobalNameSpace", Value = "False", DataType = eAttributeDataType.NonString });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Creatable", Value = "False", DataType = eAttributeDataType.NonString });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_PredeclaredId", Value = "False", DataType = eAttributeDataType.NonString });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Exposed", Value = Exposed ? "True" : "False", DataType = eAttributeDataType.NonString });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_TemplateDerived", Value = "False", DataType = eAttributeDataType.NonString });
            m.Attributes._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Customizable", Value = "False", DataType = eAttributeDataType.NonString });

            //m.Code = _project.GetBlankClassModule(Name, Exposed);
            m.Private = !Exposed;
            //m.ClassID=
            _list.Add(m);
            return m;
        }
    }
}
