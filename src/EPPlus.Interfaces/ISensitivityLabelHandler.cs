#if(!NET35)
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Interfaces.SensitivityLabels;

public interface IDecryptedPackage
{
    public MemoryStream PackageStream { get; set; }
    public object ProtectionInformation { get; set; }
}
public interface ISensitivityLabelHandler
{
    public Task InitAsync();
    public Task<IDecryptedPackage> DecryptPackageAsync(MemoryStream packageStream, string Id);
    public Task<MemoryStream> EncryptPackageAsync(IDecryptedPackage package, string Id);
    public void UpdateLabelList(IEnumerable<IExcelSensibilityLabel> list, string Id);
    public IEnumerable<IExcelSensibilityLabel> GetLabels();
}

public interface IExcelSensibilityLabel 
{
    public string Id { get;  }
    public string Name { get; }
    public string Description { get; }
    public string Tooltip { get; }
    public IExcelSensibilityLabel Parent { get; }
    public string Color { get; }
    public bool Enabled { get; }
    public bool Removed { get; }
    public string SiteId { get; }
    public eMethod Method { get; }
    public eContentBits ContentBits { get; }
}
public interface IExcelSensibilityLabelUpdate
{
    public void Update(string name, string tooltip, string description, string color, IExcelSensibilityLabel parent);
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
#endif