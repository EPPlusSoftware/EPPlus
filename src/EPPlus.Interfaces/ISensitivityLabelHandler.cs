#if(!NET35)
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Interfaces
{
    public interface ISensitivityLabelHandler
    {
        public Task Init(int Id);
        public MemoryStream EncryptPackage(MemoryStream packageStream, ref int Id);
        public Task<MemoryStream> DecryptPackage(MemoryStream packageStream, int Id);
        public IList<IExcelSensibilityLabel> UpdateLabelList(IList<IExcelSensibilityLabel> list, int Id);
        public void SetActiveLabel(string name);
    }

    public interface IExcelSensibilityLabel
    {
        public string Id { get;  }
        public string Name { get;  }
        public string Description { get;  }
        public bool Enabled { get; }
        public bool Removed { get; }
        public string SiteId { get; }
        public eMethod Method { get; }
        public eContentBits ContentBits { get; }
    }
    public enum eMethod
    {
        Empty,
        Standard,
        Privileged
    }
    [Flags]
    public enum eContentBits
    {
        None = 0,
        Header = 1,
        Footer = 2,
        Watermark = 4,
        Encryption = 8
    }
}
#endif