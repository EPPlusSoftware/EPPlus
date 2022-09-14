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
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal class V3NormalizedDataHashInputProvider : ContentHashInputProvider
    {
        public V3NormalizedDataHashInputProvider(ExcelVbaProject project) : base(project)
        {
        }

        protected override void CreateHashInputInternal(MemoryStream ms)
        {
            BinaryWriter bw = new BinaryWriter(ms);
            CreateV3NormalizedDataHashInput(bw);
        }

        private void CreateV3NormalizedDataHashInput(BinaryWriter bw)
        {
            var p = base.Project;

            //MS-OVBA 2.4.2.5 V3 Content Normalized Data
            bw.Write((ushort)1);        // PROJECTSYSKIND.Id
            bw.Write((uint)4);          // PROJECTSYSKIND.Size

            bw.Write((ushort)2);        // PROJECTLCID.Id
            bw.Write((uint)4);          // PROJECTLCID.Size
            bw.Write((uint)p.Lcid);     // PROJECTLCID.Lcid

            bw.Write((ushort)0x14);     // PROJECTLCIDINVOKE.Id
            bw.Write((uint)4);          // PROJECTLCIDINVOKE.Size
            bw.Write((uint)p.LcidInvoke); // PROJECTLCIDINVOKE.LcidInvoke

            bw.Write((ushort)3);        // PROJECTCODEPAGE.Id
            bw.Write((uint)2);          // PROJECTCODEPAGE.Size

            bw.Write((ushort)4);        // PROJECTNAME.Id
            var nameBytes = Encoding.GetEncoding(p.CodePage).GetBytes(p.Name);
            bw.Write((uint)nameBytes.Length);   // PROJECTNAME.SizeOfProjectName
            bw.Write(nameBytes);        // PROJECTNAME.ProjectName

            /*
             * APPEND Buffer WITH PROJECTDOCSTRING.Id (section 2.3.4.2.1.7) 
             * APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocString (section 2.3.4.2.1.7) of Storage 
             * APPEND Buffer WITH PROJECTDOCSTRING.Reserved (section 2.3.4.2.1.7) 
             * APPEND Buffer WITH PROJECTDOCSTRING.SizeOfDocStringUnicode (section 2.3.4.2.1.7) of Storage
             * */

            bw.Write((ushort)5);          // PROJECTDOCSTRING.Id
            
            var descriptionBytes = Encoding.GetEncoding(p.CodePage).GetBytes(p.Description);
            bw.Write((uint)descriptionBytes.Length);    // PROJECTDOCSTRING.SizeOfDocString
            
            bw.Write((ushort)0x0040);   // PROJECTDOCSTRING.Reserved
            
            var descriptionUnicodeBytes = Encoding.Unicode.GetBytes(p.Description);
            bw.Write((uint)descriptionUnicodeBytes.Length); //PROJECTDOCSTRING.SizeOfDocStringUnicode

            /*
             * APPEND Buffer WITH PROJECTHELPFILEPATH.Id (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile1 (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.Reserved (section 2.3.4.2.1.8) of Storage 
             * APPEND Buffer WITH PROJECTHELPFILEPATH.SizeOfHelpFile2 (section 2.3.4.2.1.8) of Storage
             **/
            bw.Write((ushort)6);        // PROJECTHELPFILEPATH.Id
            
            var helpFile1Bytes = Encoding.GetEncoding(p.CodePage).GetBytes(p.HelpFile1);
            bw.Write((uint)helpFile1Bytes.Length);  // PROJECTHELPFILEPATH.SizeOfHelpFile1            
            
            bw.Write((ushort)0x3D);     // PROJECTHELPFILEPATH.Reserved
            
            var helpFile2Bytes = Encoding.GetEncoding(p.CodePage).GetBytes(p.HelpFile2);
            bw.Write((uint)helpFile2Bytes.Length);  // PROJECTHELPFILEPATH.SizeOfHelpFile2

        }
    }
}
