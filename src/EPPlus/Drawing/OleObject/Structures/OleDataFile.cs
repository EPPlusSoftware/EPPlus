using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
