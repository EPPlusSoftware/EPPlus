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
            //MS-OVBA 2.4.2.5 V3 Content Normalized Data
            bw.Write((ushort)0x0001);           // PROJECTSYSKIND.Id
            bw.Write((uint)0x00000004);         // PROJECTSYSKIND.Size

            bw.Write((ushort)0x0002);           // PROJECTLCID.Id
            bw.Write((uint)0x00000004);         // PROJECTLCID.Size
            bw.Write((uint)p.Lcid);             // PROJECTLCID.Lcid

            bw.Write((ushort)0x0014);           // PROJECTLCIDINVOKE.Id
            bw.Write((uint)0x00000004);         // PROJECTLCIDINVOKE.Size
            bw.Write((uint)p.LcidInvoke);       // PROJECTLCIDINVOKE.LcidInvoke

            bw.Write((ushort)0x0003);           // PROJECTCODEPAGE.Id
            bw.Write((uint)0x00000002);         // PROJECTCODEPAGE.Size
                                                
            bw.Write((ushort)0x0004);           // PROJECTNAME.Id
            var nameBytes = encoding.GetBytes(p.Name);
            bw.Write((uint)nameBytes.Length);   // PROJECTNAME.SizeOfProjectName
            bw.Write(nameBytes);                // PROJECTNAME.ProjectName

            /*
             * APPEND Buffer WITH PROJECTDOCSTRING.Id (section 2.3.4.2.1.7) 
             * APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocString (section 2.3.4.2.1.7) of Storage 
             * APPEND Buffer WITH PROJECTDOCSTRING.Reserved (section 2.3.4.2.1.7) 
             * APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocStringUnicode (section 2.3.4.2.1.7) of Storage
             * */

            bw.Write((ushort)0x0005);           // PROJECTDOCSTRING.Id
            
            var descriptionBytes = encoding.GetBytes(p.Description);
            bw.Write((uint)descriptionBytes.Length);    // PROJECTDOCSTRING.SizeOfDocString
            
            bw.Write((ushort)0x0040);           // PROJECTDOCSTRING.Reserved
            
            var descriptionUnicodeBytes = Encoding.Unicode.GetBytes(p.Description);
            bw.Write((uint)descriptionUnicodeBytes.Length); //PROJECTDOCSTRING.SizeOfDocStringUnicode

            /*
             * APPEND Buffer WITH PROJECTHELPFILEPATH.Id (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile1 (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.Reserved (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile2 (section 2.3.4.2.1.8) of Storage
             **/
            bw.Write((ushort)0x0006);               // PROJECTHELPFILEPATH.Id
            
            var helpFile1Bytes = encoding.GetBytes(p.HelpFile1);
            bw.Write((uint)helpFile1Bytes.Length);  // PROJECTHELPFILEPATH.SizeOfHelpFile1            
            
            bw.Write((ushort)0x3D);                 // PROJECTHELPFILEPATH.Reserved
            
            var helpFile2Bytes = encoding.GetBytes(p.HelpFile2);
            bw.Write((uint)helpFile2Bytes.Length);  // PROJECTHELPFILEPATH.SizeOfHelpFile2

            /*
             * APPEND Buffer WITH PROJECTHELPCONTEXT.Id (section 2.3.4.2.1.9) of Storage
             * APPEND Buffer WITH PROJECTHELPCONTEXT.Size (section 2.3.4.2.1.9) of Storage
             **/
            //Help context id
            bw.Write((ushort)0x0007);               // Id
            bw.Write((uint)0x00000004);             // Size

            /*
             * APPEND Buffer WITH PROJECTLIBFLAGS.Id (section 2.3.4.2.1.10) of Storage 
             * APPEND Buffer WITH PROJECTLIBFLAGS.Size (section 2.3.4.2.1.10) of Storage 
             * APPEND Buffer WITH PROJECTLIBFLAGS.ProjectLibFlags (section 2.3.4.2.1.10) of Storage
             **/
            bw.Write((ushort)0x0008);               // ID
            bw.Write((uint)0x00000004);             // Size
            bw.Write((uint)0x00000000);             // ProjectLibFlags

            /*
             * APPEND Buffer WITH PROJECTVERSION.Id (section 2.3.4.2.1.11) of Storage 
             * APPEND Buffer WITH PROJECTVERSION.Reserved (section 2.3.4.2.1.11) of Storage 
             * APPEND Buffer WITH PROJECTVERSION.VersionMajor (section 2.3.4.2.1.11) of Storage 
             * APPEND Buffer WITH PROJECTVERSION.VersionMinor (section 2.3.4.2.1.11) of Storage
             **/
            bw.Write((ushort)0x0009);               // Id
            bw.Write((uint)0x00000004);             // Reserved
            bw.Write((uint)p.MajorVersion);         // Major version
            bw.Write((ushort)p.MinorVersion);       // Minor version

            /*
             * APPEND Buffer WITH PROJECTCONSTANTS.Id (section 2.3.4.2.1.12) of Storage 
             * APPEND Buffer WITH PROJECTCONSTANTS.SizeOfConstants (section 2.3.4.2.1.12) of Storage 
             * APPEND Buffer WITH PROJECTCONSTANTS.Constants (section 2.3.4.2.1.12) of Storage 
             * APPEND Buffer WITH PROJECTCONSTANTS.Reserved (section 2.3.4.2.1.12) of Storage 
             * APPEND Buffer WITH PROJECTCONSTANTS.SizeOfConstantsUnicode (section 2.3.4.2.1.12) of Storage 
             * APPEND Buffer WITH PROJECTCONSTANTS.ConstantsUnicode (section 2.3.4.2.1.12) of Storage
             **/
            bw.Write((ushort)0x000C);                                           // Id

            var constantsBytes = encoding.GetBytes(p.Constants);
            bw.Write((uint)constantsBytes.Length);                              // SizeOfConstants
            bw.Write(constantsBytes);                                           // Constants
            bw.Write((ushort)0x003C);                                           // Reserved
            var constantsUnicodeBytes = Encoding.Unicode.GetBytes(p.Constants);
            bw.Write((uint)constantsUnicodeBytes.Length);                       // SizeOfConstantsUnicode
            bw.Write(constantsUnicodeBytes);                                    // ConstantsUnicode

            /*
             * FOR EACH REFERENCE (section 2.3.4.2.2.1) IN PROJECTREFERENCES.ReferenceArray (section
             * 2.3.4.2.2) of Storage
             */
            foreach (var reference in p.References)
            {
                WriteNameReference(p, bw, reference);

                if (reference.ReferenceRecordID == 0x2F)
                {
                    WriteControlReference(p, bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x33)
                {
                    WriteOrginalReference(p, bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0D)
                {
                    WriteRegisteredReference(p, bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0E)
                {
                    WriteProjectReference(p, bw, reference);
                }
            }

            // 2.3.4.2.3 PROJECTMODULES Record
            bw.Write((ushort)0x000F); // Id
            bw.Write((uint)0x00000002); // Size

            // 2.3.4.2.3.1 PROJECTCOOKIE Record
            bw.Write((ushort)0x0013);
            bw.Write((uint)0x00000002);

            /*
             * FOR EACH Module IN ProjectModules
             */
            foreach(var module in p.Modules)
            {
                WriteModuleRecord(p, bw, module);
            }
        }

        /// <summary>
        /// 2.3.4.2.2.2 REFERENCENAME Record
        /// </summary>
        /// <param name="p"></param>
        /// <param name="bw"></param>
        /// <param name="reference"></param>
        private void WriteNameReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x0016); // Id
            var refNameBytes = Encoding.GetEncoding(p.CodePage).GetBytes(reference.Name);
            bw.Write((uint)refNameBytes.Length); // Size
            bw.Write(refNameBytes); // Name
            bw.Write((ushort)0x003E); // Reserved
            var refNameUnicodeBytes = Encoding.Unicode.GetBytes(reference.Name);
            bw.Write((uint)refNameUnicodeBytes.Length);
            bw.Write(refNameUnicodeBytes);
        }

        /// <summary>
        /// 2.3.4.2.2.4 REFERENCEORIGINAL Record
        /// </summary>
        /// <param name="p"></param>
        /// <param name="bw"></param>
        /// <param name="reference"></param>
        private void WriteOrginalReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x33);             // Id
            var libIdBytes = Encoding.GetEncoding(p.CodePage).GetBytes(reference.Libid);
            bw.Write((uint)libIdBytes.Length);  // Size
            bw.Write(libIdBytes);               // LibID
        }

        /// <summary>
        /// 2.3.4.2.2.3 REFERENCECONTROL Record
        /// </summary>
        /// <param name="p"></param>
        /// <param name="bw"></param>
        /// <param name="reference"></param>
        private void WriteControlReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x002F);                                       // Id

            var controlRef = (ExcelVbaReferenceControl)reference;

            var libIdTwiddledBytes = Encoding.GetEncoding(p.CodePage).GetBytes(controlRef.LibIdTwiddled);
            // bw.Write((uint)(4 + libIdTwiddledBytes.Length + 4 + 2));     // SizeTwiddled - Size of SizeOfLibidTwiddled, LibidTwiddled, Reserved1, and Reserved2.
            bw.Write((uint)libIdTwiddledBytes.Length);                      // SizeOfLibidTwiddled             
            bw.Write(libIdTwiddledBytes);                                   // LibIDTwiddled

            bw.Write((uint)0x00000000);                                     // Reserved1
            bw.Write((ushort)0);                                            // Reserved2
            
            if(!string.IsNullOrEmpty(controlRef.LibIdExtended))
            {
                //WriteNameReference(p, bw, reference);  //Name record again
                //controlRef.ex
            }
            bw.Write((ushort)0x30); //Reserved3

            var libIdExtendedBytes = Encoding.GetEncoding(p.CodePage).GetBytes(controlRef.LibIdExtended);
            //bw.Write((uint)(4 + libIdExternalBytes.Length + 4 + 2 + 16 + 4));    // Size of SizeOfLibidExtended, LibidExtended, Reserved4, Reserved5, OriginalTypeLib, and Cookie
            bw.Write((uint)libIdExtendedBytes.Length);                      // SizeOfLibidExtended            
            bw.Write(libIdExtendedBytes);                                   // LibIdExtended
            bw.Write((uint)0);                                              //Reserved4
            bw.Write((ushort)0);                                            //Reserved5
            bw.Write(controlRef.OriginalTypeLib.ToByteArray());             // OriginalTypeLib
            bw.Write(controlRef.Cookie);                                    //Cookie
        }

        /// <summary>
        /// 2.3.4.2.2.5 REFERENCEREGISTERED Record
        /// </summary>
        /// <param name="p"></param>
        /// <param name="bw"></param>
        /// <param name="reference"></param>
        private void WriteRegisteredReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x000D);   // Id
            var libIdBytes = Encoding.GetEncoding(p.CodePage).GetBytes(reference.Libid);
            bw.Write((uint)libIdBytes.Length); // Size of LibId
            bw.Write(libIdBytes);       // LibID            
            bw.Write((uint)0);          // Reserved1
            bw.Write((ushort)0);        // Reserved2
        }

        /// <summary>
        /// 2.3.4.2.2.6 REFERENCEPROJECT Record
        /// </summary>
        /// <param name="bw"></param>
        /// <param name="reference"></param>
        private void WriteProjectReference(ExcelVbaProject p, BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x000E); // Id
            var projRef = (ExcelVbaReferenceProject)reference;
            var libIdBytes = Encoding.GetEncoding(p.CodePage).GetBytes(projRef.Libid);
            var libIdRelativeBytes = Encoding.GetEncoding(p.CodePage).GetBytes(projRef.LibIdRelative);
            bw.Write((uint)libIdBytes.Length);  // SizeOfLibidAbsolute
            bw.Write(libIdBytes);               // LibAbsolute
            bw.Write((uint)libIdRelativeBytes.Length); // SizeOfLibIdRelative
            bw.Write(libIdRelativeBytes);       // LibIdRelative
            bw.Write(projRef.MajorVersion);
            bw.Write(projRef.MinorVersion);
        }

        private void WriteModuleRecord(ExcelVbaProject p, BinaryWriter bw, ExcelVBAModule module)
        {
            // IF Module.ModuleType.Id = 0x21 THEN
            if (module.Type == eModuleType.Module)
            {
                bw.Write((ushort)0x0021);           // Id
                bw.Write((uint)0x00000000);         // Reserved
            }
            else if(module.Type == eModuleType.Document || module.Type == eModuleType.Class || module.Type == eModuleType.Designer)
            {
                bw.Write((ushort)0x0022);           // Id
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
            foreach(var ch in text)
            {
                if((ch == 0xA || ch == 0xD) && textBuffer.Count > 0)
                {
                    lines.Add(textBuffer.ToArray());
                    textBuffer.Clear();
                }
                else
                {
                    textBuffer.Add(ch);
                }
            }

            var hashModuleNameFlag = false;
            foreach(var line in lines)
            {
                var lineText = Encoding.GetEncoding(p.CodePage).GetString(line);
                //if(string.Compare(lineText, "attribute", true) != 0)
                if(!lineText.ToLower().StartsWith("attribute"))
                {
                    hashModuleNameFlag = true;
                    bw.Write(line);
                    bw.Write((byte)'\n');
                }
                else if(lineText.StartsWith("Attribute VB_Name = "))
                {
                    continue;
                }
                else if(DefaultAttributes.Contains(lineText) == false)
                {
                    hashModuleNameFlag = true;
                    bw.Write(line);
                    bw.Write((byte)'\n');
                }
            }
            if(hashModuleNameFlag)
            {
                /*
                 * IF exist MODULENAME.ModuleNameUnicode
                 *   APPEND Buffer WITH MODULENAME.ModuleNameUnicode (section 2.3.4.2.3.2.2)
                 * ELSE IF exist MODULENAME.ModuleName
                 *   APPEND Buffer WITH MODULENAME.ModuleName (section 2.3.4.2.3.2.1)
                 * END IF
                 * APPEND Buffer WITH “\n”
                 */
                if(!string.IsNullOrEmpty(module.NameUnicode))
                {
                    var nameUnicodeBytes = Encoding.Unicode.GetBytes(module.NameUnicode);
                    bw.Write(nameUnicodeBytes);
                }
                else
                {
                    var nameBytes = Encoding.Unicode.GetBytes(module.Name);
                    bw.Write(nameBytes);
                }
                bw.Write((byte)'\n');
            }

            // APPEND Buffer WITH Terminator (section 2.3.4.2) of Storage
            bw.Write((ushort)0x0010);
            // APPEND Buffer WITH Reserved (section 2.3.4.2) of Storage
            bw.Write((uint)0x00000000);
        }


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
            foreach(var line in lines)
            {
                if(line.StartsWith("[") && line.EndsWith("]"))
                {
                    currentCategory = line.Substring(1, line.Length - 2);
                    continue;
                }
                if(currentCategory == "Host Extender Info" && !string.IsNullOrEmpty(line))
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
                    if(propertyName.ToLower() == "baseclass")
                    {
                        /*
                         * IF property is ProjectDesignerModule THEN 
                         *   APPEND Buffer WITH output of NormalizeDesignerStorage(ProjectDesignerModule) (section 2.4.2.2) 
                         * END IF
                         */
                        FormsNormalizedDataHashInputProvider.NormalizeDesigner(p, bw, propertyValue);
                    }
                    if(propertyName != "ID" && propertyName != "Document" && propertyName != "CMG" && propertyName != "DPB" && propertyName != "GC")
                    {
                        /*
                         * APPEND Buffer WITH the string “Host Extender Info” APPEND Buffer WITH HostExtenderRef without NWLN (section 2.3.1.18)
                         */
                        bw.Write(encoding.GetBytes(propertyName));
                        //propertyValue = propertyValue.Replace("\"", string.Empty);
                        bw.Write(encoding.GetBytes(propertyValue));
                    }
                }
            }
            if(hostExtenders.Count > 0)
            {
                bw.Write(encoding.GetBytes("Host Extender Info"));
            }
            foreach(var hostExtender in hostExtenders)
            {
                bw.Write(encoding.GetBytes(hostExtender));
            }
        }
    }
}
