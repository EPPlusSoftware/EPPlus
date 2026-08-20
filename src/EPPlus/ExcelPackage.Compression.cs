using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
    public sealed partial class ExcelPackage
    {
        static ExcelPackage()
        {
            CompressionStreamFactory.Provider = () => EPPlusMemoryManager.GetStream();
            CompressionStreamFactory.BufferProvider = buffer => EPPlusMemoryManager.GetStream(buffer);
        }
    }
}