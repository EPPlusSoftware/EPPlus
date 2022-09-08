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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using comTypes = System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security;

namespace OfficeOpenXml.Utils.CompundDocument
{
    internal class CompoundDocument
    {        
        internal class StoragePart
        {
            public StoragePart()
            {

            }
            internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
            internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
        }
        internal StoragePart Storage = null;
        internal CompoundDocument()
        {
            Storage = new StoragePart();
        }
        internal CompoundDocument(MemoryStream ms)
        {
            Read(ms);
        }
        internal CompoundDocument(FileInfo fi)
        {
            Read(fi);
        }

        internal static bool IsCompoundDocument(FileInfo fi)
        {
            return CompoundDocumentFile.IsCompoundDocument(fi);
        }
        internal static bool IsCompoundDocument(MemoryStream ms)
        {
            return CompoundDocumentFile.IsCompoundDocument(ms);
        }

        internal CompoundDocument(byte[] doc)
        {
            Read(doc);
        }
        internal void Read(FileInfo fi)
        {
            var b = File.ReadAllBytes(fi.FullName);
            Read(b);
        }
        internal void Read(byte[] doc) 
        {
            using (var ms = RecyclableMemory.GetStream(doc))
            {
                Read(ms);
            }
        }
        internal void Read(MemoryStream ms)
        {
            using (var doc = new CompoundDocumentFile(ms))
            {
                Storage = new StoragePart();
                GetStorageAndStreams(Storage, doc.RootItem);
            }
        }

        internal List<CompoundDocumentItem> GetDirs()
        {
            var doc = new CompoundDocumentFile();
            WriteStorageAndStreams(Storage, doc.RootItem);
            return doc.FlattenDirs();
        }

        private void GetStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach(var item in parent.Children)
            {
                if(item.ObjectType==1)      //Substorage
                {
                    var part = new StoragePart();
                    storage.SubStorage.Add(item.Name, part);
                    GetStorageAndStreams(part, item);
                }
                else if(item.ObjectType==2) //Stream
                {
                    storage.DataStreams.Add(item.Name, item.Stream);
                }
            }
        }
        internal void Save(MemoryStream ms)
        {
            var doc = new CompoundDocumentFile();
            WriteStorageAndStreams(Storage, doc.RootItem);
            doc.Write(ms);
        }

        private void WriteStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach(var item in storage.SubStorage)
            {
                var c = new CompoundDocumentItem() { Name = item.Key, ObjectType = 1, Stream = null, StreamSize = 0, Parent = parent };
                parent.Children.Add(c);
                WriteStorageAndStreams(item.Value, c);
            }
            foreach (var item in storage.DataStreams)
            {
                var c = new CompoundDocumentItem() { Name = item.Key, ObjectType = 2, Stream = item.Value, StreamSize = (item.Value == null ? 0 : item.Value.Length), Parent = parent };
                parent.Children.Add(c);
            }
        }
    }
}