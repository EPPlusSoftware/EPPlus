/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal class V3NormalizedDataHashInputProvider : ContentHashInputProvider
    {
        public V3NormalizedDataHashInputProvider(ExcelVbaProject project) : base(project)
        {
        }

        /// <summary>
        /// This list of strings is taken from [MS-OVBA] v20220517, 2.4.2.5 V3 Content Normalized Data
        /// </summary>
        private readonly List<string> DefaultAttributes = new List<string>
        {
            "Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True"
        };

        protected override void CreateHashInputInternal(MemoryStream ms)
        {
            BinaryWriter bw = new BinaryWriter(ms);
            CreateV3NormalizedDataHashInput(bw);
            NormalizeProjectStream(bw);
        }

        /// <summary>
        /// This is an implementation of the meta code described in
        /// [MS-OVBA] v20220517, 2.4.2.5 V3 Content Normalized Data
        /// </summary>
        /// <param name="bw"></param>
        private void CreateV3NormalizedDataHashInput(BinaryWriter bw)
        {
            var p = base.Project;
            var encoding = Encoding.GetEncoding(p.CodePage);

            /******************************************
             * 2.3.4.2.1.1 PROJECTSYSKIND Record      *
             ******************************************/

            // APPEND Buffer WITH PROJECTSYSKIND.Id (section 2.3.4.2.1.1) of Storage
            bw.Write((ushort)0x0001);

            // APPEND Buffer WITH PROJECTSYSKIND.Size (section 2.3.4.2.1.1) of Storage
            bw.Write((uint)0x00000004); 

            /******************************************
             * 2.3.4.2.1.4 PROJECTLCIDINVOKE Record   *
             ******************************************/

            // APPEND Buffer WITH PROJECTLCID.Id (section 2.3.4.2.1.3) of Storage
            
            bw.Write((ushort)0x0002);

            // APPEND Buffer WITH PROJECTLCID.Size (section 2.3.4.2.1.3) of Storage
            bw.Write((uint)0x00000004);

            // APPEND Buffer WITH PROJECTLCID.Lcid (section 2.3.4.2.1.3) of Storage
            bw.Write((uint)p.Lcid);

            /******************************************
             * 2.3.4.2.1.4 PROJECTLCIDINVOKE Record   *
             ******************************************/

            // APPEND Buffer WITH PROJECTLCIDINVOKE.Id (section 2.3.4.2.1.4) of Storage
            bw.Write((ushort)0x0014);

            // APPEND Buffer WITH PROJECTLCIDINVOKE.Size (section 2.3.4.2.1.4) of Storage
            bw.Write((uint)0x00000004);

            // APPEND Buffer WITH PROJECTLCIDINVOKE.LcidInvoke (section 2.3.4.2.1.4) of Storage
            bw.Write((uint)p.LcidInvoke);

            /******************************************
             * 2.3.4.2.1.5 PROJECTCODEPAGE Record     *
             ******************************************/

            // APPEND Buffer WITH PROJECTCODEPAGE.Id (section 2.3.4.2.1.5) of Storage
            bw.Write((ushort)0x0003);

            // APPEND Buffer WITH PROJECTCODEPAGE.Size (section 2.3.4.2.1.5) of Storage
            bw.Write((uint)0x00000002);

            /******************************************
             * 2.3.4.2.1.6 PROJECTNAME Record         *
             ******************************************/

            // APPEND Buffer WITH PROJECTNAME.Id (section 2.3.4.2.1.6) of Storage
            bw.Write((ushort)0x0004);

            var nameBytes = encoding.GetBytes(p.Name);

            // APPEND Buffer WITH PROJECTNAME.SizeOfProjectName (section 2.3.4.2.1.6) of Storage
            bw.Write((uint)nameBytes.Length);

            // APPEND Buffer WITH PROJECTNAME.ProjectName (section 2.3.4.2.1.6) of Storage
            bw.Write(nameBytes);

            /******************************************
             * 2.3.4.2.1.7 PROJECTDOCSTRING Record    *
             ******************************************/

            // APPEND Buffer WITH PROJECTDOCSTRING.Id (section 2.3.4.2.1.7) 
            bw.Write((ushort)0x0005);

            // APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocString (section 2.3.4.2.1.7) of Storage 
            var descriptionBytes = encoding.GetBytes(p.Description);
            bw.Write((uint)descriptionBytes.Length);

            // APPEND Buffer WITH PROJECTDOCSTRING.Reserved (section 2.3.4.2.1.7)
            bw.Write((ushort)0x0040);

            // APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocStringUnicode (section 2.3.4.2.1.7) of Storage
            var descriptionUnicodeBytes = Encoding.Unicode.GetBytes(p.Description);
            bw.Write((uint)descriptionUnicodeBytes.Length);

            /******************************************
             * 2.3.4.2.1.8 PROJECTHELPFILEPATH Record *
             ******************************************/

            // APPEND Buffer WITH PROJECTHELPFILEPATH.Id (section 2.3.4.2.1.8) of Storage
            bw.Write((ushort)0x0006);

            // APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile1 (section 2.3.4.2.1.8) of Storage
            var helpFile1Bytes = encoding.GetBytes(p.HelpFile1);
            bw.Write((uint)helpFile1Bytes.Length);  // PROJECTHELPFILEPATH.SizeOfHelpFile1            

            // APPEND Buffer WITH PROJECTHELPFILEPATH.Reserved (section 2.3.4.2.1.8) of Storage
            bw.Write((ushort)0x003D);

            // APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile2 (section 2.3.4.2.1.8) of Storage
            var helpFile2Bytes = encoding.GetBytes(p.HelpFile2);
            bw.Write((uint)helpFile2Bytes.Length);

            /******************************************
             * 2.3.4.2.1.9 PROJECTHELPCONTEXT Record  *
             ******************************************/

            // APPEND Buffer WITH PROJECTHELPCONTEXT.Id (section 2.3.4.2.1.9) of Storage
            bw.Write((ushort)0x0007);

            // APPEND Buffer WITH PROJECTHELPCONTEXT.Size (section 2.3.4.2.1.9) of Storage
            bw.Write((uint)0x00000004);

            /******************************************
             * 2.3.4.2.1.10 PROJECTLIBFLAGS Record    *
             ******************************************/

            // APPEND Buffer WITH PROJECTLIBFLAGS.Id (section 2.3.4.2.1.10) of Storage
            bw.Write((ushort)0x0008);

            // APPEND Buffer WITH PROJECTLIBFLAGS.Size (section 2.3.4.2.1.10) of Storage
            bw.Write((uint)0x00000004);

            // APPEND Buffer WITH PROJECTLIBFLAGS.ProjectLibFlags (section 2.3.4.2.1.10) of Storage
            bw.Write((uint)0x00000000);

            /******************************************
             * 2.3.4.2.1.11 PROJECTVERSION Record     *
             ******************************************/

            // APPEND Buffer WITH PROJECTVERSION.Id (section 2.3.4.2.1.11) of Storage 
            bw.Write((ushort)0x0009);

            // APPEND Buffer WITH PROJECTVERSION.Reserved (section 2.3.4.2.1.11) of Storage
            bw.Write((uint)0x00000004);

            // APPEND Buffer WITH PROJECTVERSION.VersionMajor (section 2.3.4.2.1.11) of Storage
            bw.Write((uint)p.MajorVersion);

            // APPEND Buffer WITH PROJECTVERSION.VersionMinor (section 2.3.4.2.1.11) of Storage
            bw.Write((ushort)p.MinorVersion);

            /******************************************
             * 2.3.4.2.1.12 PROJECTCONSTANTS Record   *
             ******************************************/

            // APPEND Buffer WITH PROJECTCONSTANTS.Id (section 2.3.4.2.1.12) of Storage
            bw.Write((ushort)0x000C);

            var constantsBytes = encoding.GetBytes(p.Constants);

            // APPEND Buffer WITH PROJECTCONSTANTS.SizeOfConstants (section 2.3.4.2.1.12) of Storage
            bw.Write((uint)constantsBytes.Length);

            // APPEND Buffer WITH PROJECTCONSTANTS.Constants (section 2.3.4.2.1.12) of Storage
            bw.Write(constantsBytes);

            // APPEND Buffer WITH PROJECTCONSTANTS.Reserved(section 2.3.4.2.1.12) of Storage
            bw.Write((ushort)0x003C);

            var constantsUnicodeBytes = Encoding.Unicode.GetBytes(p.Constants);

            // APPEND Buffer WITH PROJECTCONSTANTS.SizeOfConstantsUnicode (section 2.3.4.2.1.12) of Storage
            bw.Write((uint)constantsUnicodeBytes.Length);

            // APPEND Buffer WITH PROJECTCONSTANTS.ConstantsUnicode (section 2.3.4.2.1.12) of Storage
            bw.Write(constantsUnicodeBytes);

            /*
             * FOR EACH REFERENCE (section 2.3.4.2.2.1) IN PROJECTREFERENCES.ReferenceArray (section
             * 2.3.4.2.2) of Storage
             */
            foreach (var reference in p.References)
            {
                /******************************************
                 * 2.3.4.2.2.1 REFERENCE Record           *
                 ******************************************/
                HandleProjectReference(p, bw, reference);
            }

            /******************************************
             * 2.3.4.2.3 PROJECTMODULES Record        *
             ******************************************/

            // APPEND Buffer WITH PROJECTMODULES.Id (section 2.3.4.2.3) of Storage
            bw.Write((ushort)0x000F);

            // APPEND Buffer WITH PROJECTMODULES.Size (section 2.3.4.2.3) of Storage
            bw.Write((uint)0x00000002); // Size

            /******************************************
             * 2.3.4.2.3.1 PROJECTCOOKIE Record       *
             ******************************************/

            // APPEND Buffer WITH PROJECTCOOKIE.Id (section 2.3.4.2.3.1) of Storage
            bw.Write((ushort)0x0013);

            // APPEND Buffer WITH PROJECTCOOKIE.Size (section 2.3.4.2.3.1) of Storage
            bw.Write((uint)0x00000002);

            /*
             * FOR EACH Module IN ProjectModules
             */
            foreach(var module in p.Modules)
            {
                /******************************************
                 * 2.3.4.2.2.1 REFERENCE Record           *
                 ******************************************/
                WriteModuleRecord(p, bw, module);
            }

            // APPEND Buffer WITH Terminator (section 2.3.4.2) of Storage
            bw.Write((ushort)0x0010);
            // APPEND Buffer WITH Reserved (section 2.3.4.2) of Storage
            bw.Write((uint)0x00000000);

        }

        private void HandleProjectReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            var encoding = Encoding.GetEncoding(p.CodePage);

            /******************************************
             * 2.3.4.2.2.2 REFERENCENAME Record       *
             ******************************************/

            WriteNameRecord(bw, reference, encoding);

            // IF REFERENCE.ReferenceRecord.Id = 0x002F THEN
            // ELSE IF REFERENCE.ReferenceRecord.Id = 0x0033 THEN
            if (reference.ReferenceRecordID == 0x0033)
            {
                /******************************************
                 * 2.3.4.2.2.4 REFERENCEORIGINAL Record   *
                 ******************************************/

                // APPEND Buffer with REFERENCE.ReferenceOriginal.Id (section 2.3.4.2.2.4)
                bw.Write((ushort)0x33);

                var libIdBytes = encoding.GetBytes(reference.Libid);
                // APPEND Buffer with REFERENCE.ReferenceOriginal.SizeOfLibidOriginal (section 2.3.4.2.2.4)
                bw.Write((uint)libIdBytes.Length);

                // APPEND Buffer with REFERENCE.ReferenceOriginal.LibidOriginal (section 2.3.4.2.2.4)
                bw.Write(libIdBytes);

                if (reference.SecondaryReferenceRecordID == 0x002F)
                {
                    /******************************************
                     * 2.3.4.2.2.3 REFERENCECONTROL Record    *
                     ******************************************/

                    // APPEND Buffer with REFERENCE.ReferenceControl.Id (section 2.3.4.2.2.3)
                    bw.Write((ushort)0x002F);

                    var controlRef = (ExcelVbaReferenceControl)reference;
                    var libIdTwiddledBytes = encoding.GetBytes(controlRef.LibIdTwiddled);

                    // APPEND Buffer with REFERENCE.ReferenceControl.SizeOfLibidTwiddled (section 2.3.4.2.2.3)
                    bw.Write((uint)libIdTwiddledBytes.Length);

                    // APPEND Buffer with REFERENCE.ReferenceControl.LibidTwiddled (section 2.3.4.2.2.3)
                    bw.Write(libIdTwiddledBytes);

                    // APPEND Buffer with REFERENCE.ReferenceControl.Reserved1 (section 2.3.4.2.2.3)
                    bw.Write((uint)0x00000000);

                    // APPEND Buffer with REFERENCE.ReferenceControl.Reserved2 (section 2.3.4.2.2.3)
                    bw.Write((ushort)0x000);

                    //Write name record again.
                    WriteNameRecord(bw, reference, encoding);

                    /******************************************
                     * 2.3.4.2.2.3 REFERENCECONTROL Record    *
                     ******************************************/

                    // APPEND Buffer with REFERENCE.ReferenceControl.Reserved3 (section 2.3.4.2.2.3)
                    bw.Write((ushort)0x0030);

                    var libIdExtendedBytes = encoding.GetBytes(controlRef.LibIdExtended);
                    // APPEND Buffer with REFERENCE.ReferenceControl.SizeOfLibidExtended (section 2.3.4.2.2.3)           
                    bw.Write((uint)libIdExtendedBytes.Length);

                    // APPEND Buffer with REFERENCE.ReferenceControl.LibidExtended (section 2.3.4.2.2.3)
                    bw.Write(libIdExtendedBytes);

                    // APPEND Buffer with REFERENCE.ReferenceControl.Reserved4 (section 2.3.4.2.2.3)
                    bw.Write((uint)0x00000000);

                    // APPEND Buffer with REFERENCE.ReferenceControl.Reserved5(section 2.3.4.2.2.3)
                    bw.Write((ushort)0x0000);

                    // APPEND Buffer with REFERENCE.ReferenceControl.OriginalTypeLib (section 2.3.4.2.2.3)
                    bw.Write(controlRef.OriginalTypeLib.ToByteArray());

                    // APPEND Buffer with REFERENCE.ReferenceControl.Cookie (section 2.3.4.2.2.3)
                    bw.Write(controlRef.Cookie);
                }
            }
            // ELSE IF REFERENCE.ReferenceRecord.Id = 0x000D THEN
            else if (reference.ReferenceRecordID == 0x000D)
            {
                /******************************************
                 * 2.3.4.2.2.5 REFERENCEREGISTERED Record *
                 ******************************************/

                // APPEND Buffer with REFERENCE.ReferenceRegistered.Id (section 2.3.4.2.2.5)
                bw.Write((ushort)0x000D);

                var libIdBytes = Encoding.Unicode.GetBytes(reference.Libid);
                // APPEND Buffer with REFERENCE.ReferenceRegistered.SizeOfLibid (section 2.3.4.2.2.5)
                bw.Write((uint)reference.Libid.Length);

                // APPEND Buffer with REFERENCE.ReferenceRegistered.Libid converted to wide char (section 2.3.4.2.2.5)
                bw.Write(libIdBytes);

                // APPEND Buffer with REFERENCE.ReferenceRegistered.Reserved1 (section 2.3.4.2.2.5)
                bw.Write((uint)0x00000000);

                // APPEND Buffer with REFERENCE.ReferenceRegistered.Reserved2 (section 2.3.4.2.2.5)
                bw.Write((ushort)0x0000);
            }
            // ELSE IF REFERENCE.ReferenceRecord.Id = 0x000E THEN
            else if (reference.ReferenceRecordID == 0x000E)
            {
                /******************************************
                 * 2.3.4.2.2.6 REFERENCEPROJECT Record    *
                 ******************************************/

                var projRef = (ExcelVbaReferenceProject)reference;
                var libIdBytes = Encoding.GetEncoding(p.CodePage).GetBytes(projRef.Libid);
                var libIdRelativeBytes = Encoding.GetEncoding(p.CodePage).GetBytes(projRef.LibIdRelative);

                // APPEND Buffer with REFERENCE.ReferenceProject.Id (section 2.3.4.2.2.6)
                bw.Write((ushort)0x000E);

                // APPEND Buffer with REFERENCE.ReferenceProject.SizeOfLibidAbsolute(section 2.3.4.2.2.6)
                bw.Write((uint)libIdBytes.Length);

                // APPEND Buffer with REFERENCE.ReferenceProject.LibidAbsolute (section 2.3.4.2.2.6)
                bw.Write(libIdBytes);

                // APPEND Buffer with REFERENCE.ReferenceProject.SizeOfLibidRelative (section 2.3.4.2.2.6)
                bw.Write((uint)libIdRelativeBytes.Length);

                // APPEND Buffer with REFERENCE.ReferenceProject.LibidRelative(section 2.3.4.2.2.6)
                bw.Write(libIdRelativeBytes);

                // APPEND Buffer with REFERENCE.ReferenceProject.MajorVersion(section 2.3.4.2.2.6)
                bw.Write(p.MajorVersion);

                // APPEND Buffer with REFERENCE.ReferenceProject.MinorVersion (section 2.3.4.2.2.6)
                bw.Write(p.MinorVersion);
            }
        }

        private static void WriteNameRecord(BinaryWriter bw, ExcelVbaReference reference, Encoding encoding)
        {
            // APPEND Buffer WITH REFERENCENAME.Id (section 2.3.4.2.2.2)
            bw.Write((ushort)0x0016); // Id

            var refNameBytes = encoding.GetBytes(reference.Name);

            // APPEND Buffer WITH REFERENCENAME.SizeOfName (section 2.3.4.2.2.2)
            bw.Write((uint)refNameBytes.Length); // Size

            // APPEND Buffer WITH REFERENCENAME.Name(section 2.3.4.2.2.2)
            bw.Write(refNameBytes); // Name

            // APPEND Buffer WITH REFERENCENAME.Reserved (section 2.3.4.2.2.2)
            bw.Write((ushort)0x003E); // Reserved

            var refNameUnicodeBytes = Encoding.Unicode.GetBytes(reference.Name);

            // APPEND Buffer WITH REFERENCENAME.SizeOfNameUnicode (section 2.3.4.2.2.2)
            bw.Write((uint)refNameUnicodeBytes.Length);

            // APPEND Buffer WITH REFERENCENAME.NameUnicode (section 2.3.4.2.2.2)
            bw.Write(refNameUnicodeBytes);
        }

        private void WriteModuleRecord(ExcelVbaProject p, BinaryWriter bw, ExcelVBAModule module)
        {
            // IF Module.ModuleType.Id = 0x21 THEN
            if (module.Type == eModuleType.Module)
            {
                bw.Write((ushort)0x0021);           // Id
                bw.Write((uint)0x00000000);         // Reserved
            }

            // 2.3.4.2.3.2.9 MODULEREADONLY Record
            if (module.ReadOnly)
            {
                bw.Write((ushort)0x0025);           // Id
                bw.Write((uint)0x00000000);         // Reserved
            }
            // 2.3.4.2.3.2.10 MODULEPRIVATE Record
            if (module.Private)
            {
                bw.Write((ushort)0x0028);           // Id
                bw.Write((uint)0x00000000);         // Reserved
            }
            /*
             * DEFINE CompressedContainer AS array of bytes
             * DEFINE Text AS array of bytes
             * SET CompressedContainer TO ModuleStream.CompressedSourceCode
             * SET Text TO result of Decompression(CompressedContainer) (section 2.4.1)
             **/
            var vbaStorage = p.Document.Storage.SubStorage["VBA"];
            var stream = vbaStorage.DataStreams[module.Name];
            var text = VBACompression.DecompressPart(stream);
            var totalText = Encoding.GetEncoding(p.CodePage).GetString(text);

            var lines = new List<byte[]>();
            var textBuffer = new List<byte>();
            byte pc = 0x0;
            foreach(var ch in text)
            {
                if(ch == 0xA || ch == 0xD)
                {
                    if (pc == 0xD)
                    {
                        lines.Add(textBuffer.ToArray());
                        textBuffer.Clear();
                    }
                }
                else
                {
                    textBuffer.Add(ch);
                }
                pc = ch;
            }

            var hashModuleNameFlag = false;
            //var endOfLine = Encoding.GetEncoding(p.CodePage).GetBytes("\\n");
            byte endOfLine = 0xA;
            foreach (var line in lines)
            {
                var lineText = Encoding.GetEncoding(p.CodePage).GetString(line);
                /*
                 * IF Line NOT start with “attribute” when ignoring case THEN 
                 *    SET HashModuleNameFlag TO true 
                 *    APPEND Buffer WITH Line 
                 *    APPEND Buffer WITH “\n”
                 */
                if (!lineText.ToLower().StartsWith("attribute"))
                {
                    hashModuleNameFlag = true;
                    bw.Write(line);
                    bw.Write(endOfLine);
                }
                /*
                 * ELSE IF Line starts with “Attribute VB_Name = ” when ignoring case THEN
                 *    CONTINUE
                 */
                else if (lineText.StartsWith("Attribute VB_Name = "))
                {
                    continue;
                }
                /*
                 * ELSE IF Line not same with any one of DefaultAttributes THEN 
                 *    SET HashModuleNameFlag TO true 
                 *    APPEND Buffer WITH Line 
                 *    APPEND Buffer WITH “\n”
                 */
                else if(DefaultAttributes.Contains(lineText) == false)
                {
                    hashModuleNameFlag = true;
                    bw.Write(line);
                    bw.Write(endOfLine);
                }
            }
            // IF HashModuleNameFlag IS true
            if (hashModuleNameFlag)
            {
                /*
                 * IF exist MODULENAME.ModuleNameUnicode
                 *   APPEND Buffer WITH MODULENAME.ModuleNameUnicode (section 2.3.4.2.3.2.2)
                 * ELSE IF exist MODULENAME.ModuleName
                 *   APPEND Buffer WITH MODULENAME.ModuleName (section 2.3.4.2.3.2.1)
                 * END IF
                 * APPEND Buffer WITH “\n”
                 */
                if (!string.IsNullOrEmpty(module.NameUnicode))
                {
                    var nameUnicodeBytes = Encoding.Unicode.GetBytes(module.NameUnicode);
                    bw.Write(nameUnicodeBytes);
                }
                else if(!string.IsNullOrEmpty(module.Name))
                {
                    var nameBytes = Encoding.GetEncoding(p.CodePage).GetBytes(module.Name);
                    bw.Write(nameBytes);
                }
                bw.Write(endOfLine);
            }
       }

        const string HostExtenderInfo = "Host Extender Info";

        /// <summary>
        /// See 2.4.2.6 Project Normalized Data for meta code
        /// </summary>
        /// <param name="bw"></param>
        private void NormalizeProjectStream(BinaryWriter bw)
        {
            var p = base.Project;
            if(string.IsNullOrEmpty(p.ProjectStreamText))
            {
                return;
            }
            var encoding = Encoding.GetEncoding(p.CodePage);
            var lines = Regex.Split(p.ProjectStreamText, "\r\n");
            var currentCategory = string.Empty;
            var hostExtenders = new List<string>();

            /***************************************************************
             * For definition of properties, see 2.3.1.1 ProjectProperties *
             ***************************************************************/
            foreach(var line in lines)
            {
                if(line.StartsWith("[") && line.EndsWith("]"))
                {
                    currentCategory = line.Substring(1, line.Length - 2);
                    continue;
                }
                if(currentCategory == HostExtenderInfo && !string.IsNullOrEmpty(line))
                {
                    hostExtenders.Add(line);
                    continue;
                }
                else if(!string.IsNullOrEmpty(currentCategory))
                {
                    continue;
                }
                if(!string.IsNullOrEmpty(line) && line.Contains("="))
                {
                    var propertyName = line.Split('=')[0];
                    var propertyValue = line.Split('=')[1];
                    /*
                     * IF property is ProjectDesignerModule THEN 
                     *   APPEND Buffer WITH output of NormalizeDesignerStorage(ProjectDesignerModule) (section 2.4.2.2) 
                     * END IF
                     */
                    if (propertyName.Equals("baseclass", StringComparison.InvariantCultureIgnoreCase))
                    {
                        FormsNormalizedDataHashInputProvider.NormalizeDesigner(p, bw, propertyValue);
                    }
                    /*
                     * IF property NOT is ProjectId (section 2.3.1.2) 
                     * OR ProjectDocModule (section 2.3.1.4) 
                     * OR ProjectProtectionState (section 2.3.1.15) 
                     * OR ProjectPassword (section 2.3.1.16) 
                     * OR ProjectVisibilityState (section 2.3.1.17) 
                     * THEN
                     **/
                    if(propertyName != "ID" && propertyName != "Document" && propertyName != "CMG" && propertyName != "DPB" && propertyName != "GC")
                    {
                        if (propertyValue.StartsWith("\"")) propertyValue = propertyValue.Substring(1, propertyValue.Length - 2);   //Remove leading and trailing double-quotes
                        bw.Write(encoding.GetBytes(propertyName));
                        bw.Write(encoding.GetBytes(propertyValue));
                    }
                }
            }
            /*
             * IF exist string “[Host Extender Info]” THEN 
             *    APPEND Buffer WITH the string “Host Extender Info” 
             *    APPEND Buffer WITH HostExtenderRef without NWLN (section 2.3.1.18) 
             * END IF
             */
            if (hostExtenders.Count > 0)
            {
                bw.Write(encoding.GetBytes(HostExtenderInfo));
            }
            foreach(var hostExtender in hostExtenders)
            {
                bw.Write(encoding.GetBytes(hostExtender));
            }
        }
    }
}
