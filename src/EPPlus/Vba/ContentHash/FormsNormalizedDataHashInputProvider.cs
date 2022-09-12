using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal class FormsNormalizedDataHashInputProvider : ContentHashInputProvider
    {
        public FormsNormalizedDataHashInputProvider(ExcelVbaProject project) : base(project)
        {
        }

        protected override void CreateHashInputInternal(MemoryStream ms)
        {
            //MS-OVBA 2.4.2.2
            BinaryWriter bw = new BinaryWriter(ms);
            FormsNormaizedData(bw);
        }

        private void FormsNormaizedData(BinaryWriter bw)
        {
            var p = base.Project;
            var designers = GetDesignersSorted(p);
            var list=new List<SortItem>();
            foreach (var designer in designers)
            {
                var storage = p.Document.Storage.SubStorage[designer];
                NormalizeStorage(storage, list);
            }
        }

        private void NormalizeStorage(CompoundDocument.StoragePart storage, List<SortItem> list)
        {
            var children = GetSortedChildren(storage);
            foreach (var child in children)
            {
                if (child.IsStream==false)
                {
                    NormalizeStorage(storage.SubStorage[child.Name], list);
                }
                list.Add(child);
            }

        }

        private static void WriteStreamData(BinaryWriter bw, byte[] b)
        {
            int streamLength;
            if (b != null)
            {
                bw.Write(b);
                streamLength = b.Length;
            }
            else
            {
                streamLength = 0;
            }
            var zeros = 1023-(streamLength % 1023);
            for (int i = 0; i < zeros; i++)
                bw.Write((byte)0);
        }

        private class SortItem
        {
            public SortItem(string name, bool isStream)
            {
                Name = name;
                IsStream = isStream;
            }
    
            public string Name { get; set; }
            public bool IsStream { get; set; }
        }
        private IList<SortItem> GetSortedChildren(CompoundDocument.StoragePart storage)
        {
            var list = new List<SortItem>();
            list.AddRange(storage.DataStreams.Keys.Select(x => new SortItem(x, true)));
            list.AddRange(storage.SubStorage.Keys.Select(x => new SortItem(x, false)));

            list.Sort((a, b) =>
            {
                if (a.Name.Length < b.Name.Length)
                {
                    return -1;
                }
                else if (a.Name.Length > b.Name.Length)
                {
                    return 1;
                }
                var n1 = a.Name.ToUpperInvariant();
                var n2 = b.Name.ToUpperInvariant();
                for (int i = 0; i < n1.Length; i++)
                {
                    if (n1[i] < n2[i])
                    {
                        return -1;
                    }
                    else if (n1[i] > n2[i])
                    {
                        return 1;
                    }
                }
                return 0;
            });
            return list;
        }

        private static IList<string> GetDesignersSorted(ExcelVbaProject p)
        {
            var designerModules = p.Modules.Where(x => x.Type == eModuleType.Designer).Select(x => x.streamName);
            var dl = designerModules.ToList();
            dl.Sort((a, b) =>
            {
                if (a.Length < b.Length)
                {
                    return -1;
                }
                else if (a.Length > b.Length)
                {
                    return 1;
                }
                var n1 = a.ToUpperInvariant();
                var n2 = b.ToUpperInvariant();
                for (int i = 0; i < n1.Length; i++)
                {
                    if (n1[i] < n2[i])
                    {
                        return -1;
                    }
                    else if (n1[i] > n2[i])
                    {
                        return 1;
                    }
                }
                return 0;
            });
            return dl;
        }

        private void NormalizeDesignerStorage(ExcelVBAModule designerModule, BinaryWriter bw)
        {
            var buffer = new System.IO.BufferedStream(bw.BaseStream, 1023);
            //var ds = p.Document.Storage.SubStorage[designerModule.streamName];
        }
    }
}
