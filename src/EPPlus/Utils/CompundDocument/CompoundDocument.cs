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
using System.Collections.Generic;
using System.IO;

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
            internal Dictionary<string, CompoundDocumentItem> DataStreams = new Dictionary<string, CompoundDocumentItem>();
        }
        /// <summary>
        /// The root storage part of the compound document.
        /// </summary>
        internal StoragePart Storage = null;

        
        internal CompoundDocumentItem RootItem { get; private set; }

        /// <summary>
        /// Directories in the order they are saved.
        /// </summary>
        internal List<CompoundDocumentItem> Directories { get; private set; }
        internal CompoundDocument()
        {
            Storage = new StoragePart();
            RootItem = new CompoundDocumentItem("Root Entry", null, 5);
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
                Directories = doc.Directories;
            }
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
                    storage.DataStreams.Add(item.Name, item);
                }
            }
        }
        internal void Save(MemoryStream ms)
        {
            var doc = new CompoundDocumentFile(RootItem);
            WriteStorageAndStreams(Storage, RootItem);
            Directories = doc.Directories;
            doc.Write(ms);
        }
        private void WriteStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach(var item in storage.SubStorage)
            {
                var c = new CompoundDocumentItem(item.Key,null, 1, parent);
                parent.Children.Add(c);
                WriteStorageAndStreams(item.Value, c);
            }

            foreach (var item in storage.DataStreams)
            {
                parent.Children.Add(item.Value);
            }            
        }
    }
}