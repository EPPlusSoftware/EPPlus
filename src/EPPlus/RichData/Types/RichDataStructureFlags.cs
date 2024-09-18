namespace OfficeOpenXml.RichData
{
    internal enum RichDataStructureFlags
    {
        ErrorWithSubType = 0x01,
        ErrorSpill = 0x02,
        ErrorPropagated = 0x04,
        LocalImage = 0x08,
        LocalImageWithAltText = 0x10
    }
}