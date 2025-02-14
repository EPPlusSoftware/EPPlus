/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System.IO;

namespace OfficeOpenXml.Drawing.OleObject.Structures
{
    internal static class OleDataFile
    {
        internal const string CONTENTS_STREAM_NAME = "CONTENTS";
        internal const string EMBEDDEDODF_STREAM_NAME = "EmbeddedOdf";

        internal static void CreateDataFileDataStream(CompoundDocument _document, string streamName, byte[] fileData)
        {
            _document.Storage.DataStreams.Add(streamName, new CompoundDocumentItem(streamName, fileData));
        }

        internal static void CreateDataFileObject(OleObjectDataStructures _oleDataStructures, byte[] fileData)
        {
            _oleDataStructures.DataFile = fileData;
        }

        internal static int ReadDataFileObject(OleObjectDataStructures _oleObjectDataStructures, byte[] fileData)
        {
            using (var ms = RecyclableMemory.GetStream(fileData))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleObjectDataStructures.DataFile = new byte[br.BaseStream.Length];
                return br.Read(_oleObjectDataStructures.DataFile, 0, (int)br.BaseStream.Length - 1);
            }
        }
    }
}
