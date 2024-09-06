using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Interfaces
{
    public interface ISensitivityLabelHandler
    {
        public MemoryStream EncryptPackage(Stream packageStream);
        public MemoryStream DecryptPackage(Stream packageStream);
        public void UpdateLabels(IList<IExcelSensibilityLabel> list);
        public IList<IExcelSensibilityLabel> GetLabels();
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
