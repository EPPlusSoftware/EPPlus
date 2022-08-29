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
using System.Xml;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Compatibility;
using System.Text;
using OfficeOpenXml.Packaging;
using System.Diagnostics;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Configuration;
#if (Core)
using Microsoft.Extensions.Configuration;
#endif
namespace OfficeOpenXml
{
    /// <summary>
    /// Represents an Excel XLSX file package.  
    /// This is the top-level object to access all parts of the document.
    /// </summary>
    /// <remarks>
    /// <example>
    /// <code>
	///     FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
	/// 	if (newFile.Exists)
	/// 	{
	/// 		newFile.Delete();  // ensures we create a new workbook
	/// 		newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
	/// 	}
	/// 	using (ExcelPackage package = new ExcelPackage(newFile))
    ///     {
    ///         // add a new worksheet to the empty workbook
    ///         ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
    ///         //Add the headers
    ///         worksheet.Cells[1, 1].Value = "ID";
    ///         worksheet.Cells[1, 2].Value = "Product";
    ///         worksheet.Cells[1, 3].Value = "Quantity";
    ///         worksheet.Cells[1, 4].Value = "Price";
    ///         worksheet.Cells[1, 5].Value = "Value";
    ///
    ///         //Add some items...
    ///         worksheet.Cells["A2"].Value = "12001";
    ///         worksheet.Cells["B2"].Value = "Nails";
    ///         worksheet.Cells["C2"].Value = 37;
    ///         worksheet.Cells["D2"].Value = 3.99;
    ///
    ///         worksheet.Cells["A3"].Value = "12002";
    ///         worksheet.Cells["B3"].Value = "Hammer";
    ///         worksheet.Cells["C3"].Value = 5;
    ///         worksheet.Cells["D3"].Value = 12.10;
    ///
    ///         worksheet.Cells["A4"].Value = "12003";
    ///         worksheet.Cells["B4"].Value = "Saw";
    ///         worksheet.Cells["C4"].Value = 12;
    ///         worksheet.Cells["D4"].Value = 15.37;
    ///
    ///         //Add a formula for the value-column
    ///         worksheet.Cells["E2:E4"].Formula = "C2*D2";
    ///
    ///            //Ok now format the values;
    ///         using (var range = worksheet.Cells[1, 1, 1, 5]) 
    ///          {
    ///             range.Style.Font.Bold = true;
    ///             range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    ///             range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
    ///             range.Style.Font.Color.SetColor(Color.White);
    ///         }
    ///
    ///         worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
    ///         worksheet.Cells["A5:E5"].Style.Font.Bold = true;
    ///
    ///         worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2,3,4,3).Address);
    ///         worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
    ///         worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";
    ///
    ///         //Create an autofilter for the range
    ///         worksheet.Cells["A1:E4"].AutoFilter = true;
    ///
    ///         worksheet.Cells["A1:E5"].AutoFitColumns(0);
    ///
    ///         // lets set the header text 
    ///         worksheet.HeaderFooter.oddHeader.CenteredText = "&amp;24&amp;U&amp;\"Arial,Regular Bold\" Inventory";
    ///         // add the page number to the footer plus the total number of pages
    ///         worksheet.HeaderFooter.oddFooter.RightAlignedText =
    ///         string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
    ///         // add the sheet name to the footer
    ///         worksheet.HeaderFooter.oddFooter.CenteredText = ExcelHeaderFooter.SheetName;
    ///         // add the file path to the footer
    ///         worksheet.HeaderFooter.oddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
    ///
    ///         worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
    ///         worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];
    ///
    ///          // Change the sheet view to show it in page layout mode
    ///           worksheet.View.PageLayoutView = true;
    ///
    ///         // set some document properties
    ///         package.Workbook.Properties.Title = "Invertory";
    ///         package.Workbook.Properties.Author = "Jan Källman";
    ///         package.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel 2007 workbook using EPPlus";
    ///
    ///         // set some extended property values
    ///         package.Workbook.Properties.Company = "AdventureWorks Inc.";
    ///
    ///         // set some custom property values
    ///         package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
    ///         package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
    ///
    ///         // save our new workbook and we are done!
    ///         package.Save();
    ///
    ///       }
    ///
    ///       return newFile.FullName;
    /// </code>
    /// More samples can be found at  <a href="https://github.com/EPPlusSoftware/EPPlus/">https://github.com/EPPlusSoftware/EPPlus/</a>
    /// </example>
    /// </remarks>
	public sealed partial class ExcelPackage : IDisposable
	{
        internal bool _isDisposed = false;
        internal const bool preserveWhitespace=false;
        Stream _stream = null;
        private bool _isExternalStream=false;
        internal ExcelPackage _loadedPackage = null;
		#region Properties
		/// <summary>
		/// Extention Schema types
		/// </summary>
        internal const string schemaXmlExtension = "application/xml";
        internal const string schemaRelsExtension = "application/vnd.openxmlformats-package.relationships+xml";
        /// <summary>
		/// Main Xml schema name
		/// </summary>
        internal const string schemaWorkbook = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        internal const string schemaMain = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                                            
		/// <summary>
		/// Relationship schema name
		/// </summary>
		internal const string schemaRelationships = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                                                                              
        internal const string schemaDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string schemaSheetDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        internal const string schemaMarkupCompatibility = @"http://schemas.openxmlformats.org/markup-compatibility/2006";
        internal const string schemaChart14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";

        internal const string schemaMicrosoftVml = @"urn:schemas-microsoft-com:vml";
        internal const string schemaMicrosoftOffice = "urn:schemas-microsoft-com:office:office";
        internal const string schemaMicrosoftExcel = "urn:schemas-microsoft-com:office:excel";
        
        internal const string schemaChart = @"http://schemas.openxmlformats.org/drawingml/2006/chart";                                                        
        internal const string schemaHyperlink = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        internal const string schemaComment = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        internal const string schemaImage = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        internal const string schemaThemeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        internal const string schemaChartStyle = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

        //Chart styling
        internal const string schemaChartStyleRelationships = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
        internal const string schemaChartColorStyleRelationships = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

        internal const string schemaThemeOverrideRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";

        //Office properties
        internal const string schemaCore = @"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        internal const string schemaExtended = @"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        internal const string schemaCustom = @"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        internal const string schemaDc = @"http://purl.org/dc/elements/1.1/";
        internal const string schemaDcTerms = @"http://purl.org/dc/terms/";
        internal const string schemaDcmiType = @"http://purl.org/dc/dcmitype/";
        internal const string schemaXsi = @"http://www.w3.org/2001/XMLSchema-instance";
        internal const string schemaVt = @"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        
        internal const string schemaMainX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        internal const string schemaMainX15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
        internal const string schemaMainXm = "http://schemas.microsoft.com/office/excel/2006/main";
        internal const string schemaXr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
        internal const string schemaXr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";

        //Chart Ex
        internal const string schemaMc2006 = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        internal const string schemaChartExMain = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        internal const string schemaChartEx2015_9_8 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
        internal const string schemaChartEx2015_10_21 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";
        internal const string schemaChartEx2016_5_10 = "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";
        internal const string schemaChartExRelationships = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

        internal const string schemaSlicer = "http://schemas.microsoft.com/office/drawing/2012/slicer";
        internal const string schemaDrawings2010 = "http://schemas.microsoft.com/office/drawing/2010/main";
        internal const string schemaSlicer2010 = "http://schemas.microsoft.com/office/drawing/2010/slicer";
        internal const string schemaRelationshipsSlicer = "http://schemas.microsoft.com/office/2007/relationships/slicer";
        internal const string schemaRelationshipsSlicerCache = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
        //Threaded comments
        internal const string schemaThreadedComments = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
        internal const string schemaThreadedComment = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
        //Persons
        internal const string schemaPersonsRelationShips = "http://schemas.microsoft.com/office/2017/10/relationships/person";

        // Richdata (used in worksheet.sortstate)
        internal const string schemaRichData2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";

        //Package reference
        private Packaging.ZipPackage _zipPackage;
		internal ExcelWorkbook _workbook;
        /// <summary>
        /// Maximum number of columns in a worksheet (16384). 
        /// </summary>
        public const int MaxColumns = 16384;
        /// <summary>
        /// Maximum number of rows in a worksheet (1048576). 
        /// </summary>
        public const int MaxRows = 1048576;
        internal readonly List<Action> BeforeSave=new List<Action>();
        private static ExcelPackageConfiguration _configuration = new ExcelPackageConfiguration();
        internal static ExcelPackageConfiguration GlobalConfiguration => _configuration;
        private List<ExcelInitializationError> _initErrors = new List<ExcelInitializationError>();
        #endregion
        #region ExcelPackage Constructors
        /// <summary>
        /// Create a new instance of the ExcelPackage. 
        /// Output is accessed through the Stream property, using the <see cref="SaveAs(FileInfo)"/> method or later set the <see cref="File" /> property.
        /// </summary>
        public ExcelPackage()
        {
            Init();
            ConstructNewFile(null);
        }
        /// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing file or creates a new file. 
		/// </summary>
		/// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        public ExcelPackage(FileInfo newFile)
		{
            Init();
            File = newFile;
            ConstructNewFile(null);
        }
        /// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing file or creates a new file. 
		/// </summary>
		/// <param name="path">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        public ExcelPackage(string path)
            : this(new FileInfo(path))
        { }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing file or creates a new file. 
        /// </summary>
        /// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        /// <param name="password">Password for an encrypted package</param>
        public ExcelPackage(FileInfo newFile, string password)
        {
            Init();
            File = newFile;
            ConstructNewFile(password);
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing file or creates a new file. 
        /// </summary>
        /// <param name="path">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        /// <param name="password">Password for an encrypted package</param>
        public ExcelPackage(string path, string password)
            : this(new FileInfo(path), password)
        { }
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing template.
		/// If newFile exists, it will be overwritten when the Save method is called
		/// </summary>
		/// <param name="newFile">The name of the Excel file to be created</param>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		public ExcelPackage(FileInfo newFile, FileInfo template)
		{
            Init();
            File = newFile;
            CreateFromTemplate(template, null);
		}
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing template.
        /// If newFile exists, it will be overwritten when the Save method is called
        /// </summary>
        /// <param name="newFile">The name of the Excel file to be created</param>
        /// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="password">Password to decrypted the template</param>
        public ExcelPackage(FileInfo newFile, FileInfo template, string password)
        {
            Init();
            File = newFile;
            CreateFromTemplate(template, password);
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing template.
        /// If newFile exists, it will be overwritten when the Save method is called
        /// </summary>
        /// <param name="newFilePath">The name of the Excel file to be created</param>
        /// <param name="templatePath">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="password">Password to decrypted the template</param>
        public ExcelPackage(string newFilePath, string templatePath, string password)
            : this(new FileInfo(newFilePath), new FileInfo(templatePath), password)
        { }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing template.
        /// </summary>
        /// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="useStream">if true use a stream. If false create a file in the temp dir with a random name</param>
        public ExcelPackage(FileInfo template, bool useStream)
        {
            Init();
            CreateFromTemplate(template, null);
            if (useStream == false)
            {
                File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
            }
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a existing template.
        /// </summary>
        /// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="useStream">if true use a stream. If false create a file in the temp dir with a random name</param>
        /// <param name="password">Password to decrypted the template</param>
        public ExcelPackage(FileInfo template, bool useStream, string password)
        {
            Init();
            CreateFromTemplate(template, password);
            if (useStream == false)
            {
                File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
            }
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
        public ExcelPackage(Stream newStream) 
        {
            Init();
            if (newStream.CanSeek && newStream.Length == 0)
            {
                _stream = newStream;
                _isExternalStream = true;
                ConstructNewFile(null);
            }
            else
            {                
                Load(newStream);
            }
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
        /// <param name="Password">The password to decrypt the document</param>
        public ExcelPackage(Stream newStream, string Password)
        {
            if (!(newStream.CanRead && newStream.CanWrite))
            {
                throw new Exception("The stream must be read/write");
            }

            Init();
            if (newStream.Length > 0)
            {
                Load(newStream,Password);
            }
            else
            {
                _stream = newStream;
                _isExternalStream = true;
                _zipPackage = new Packaging.ZipPackage(_stream);
                CreateBlankWb();
            }
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
        /// <param name="templateStream">This stream is copied to the output stream at load</param>
        public ExcelPackage(Stream newStream, Stream templateStream)
        {
            if (newStream.CanSeek && newStream.Length > 0)
            {
                throw(new Exception("The output stream must be empty. Length > 0"));
            }
            else if (!(newStream.CanRead && newStream.CanWrite))
            {
                throw new Exception("The stream must be read/write");
            }
            Init();
            Load(templateStream, newStream, null);
        }
        /// <summary>
        /// Create a new instance of the ExcelPackage class based on a stream
        /// </summary>
        /// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
        /// <param name="templateStream">This stream is copied to the output stream at load</param>
        /// <param name="Password">Password to decrypted the template</param>
        public ExcelPackage(Stream newStream, Stream templateStream, string Password)
        {
            if (newStream.Length > 0)
            {
                throw (new Exception("The output stream must be empty. Length > 0"));
            }
            else if (!(newStream.CanRead && newStream.CanWrite))
            {
                throw new Exception("The stream must be read/write");
            }
            Init();
            Load(templateStream, newStream, Password);
        }
        #endregion
        /// <summary>
        /// Init values here
        /// </summary>
        private void Init()
        {
            DoAdjustDrawings = true;
#if (Core)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);  //Add Support for codepage 1252

            var  isWorksheets1Based = ExcelConfigurationReader.GetJsonConfigValue("EPPlus:ExcelPackage:Compatibility:IsWorksheets1Based", _configuration, _initErrors);

#else
            var isWorksheets1Based = ExcelConfigurationReader.GetValueFromAppSettings("EPPlus:ExcelPackage.Compatibility.IsWorksheets1Based", _configuration, _initErrors);
#endif
            if (isWorksheets1Based != null)
            {
                if(Boolean.TryParse(isWorksheets1Based.ToLowerInvariant(), out bool value))
                {
                    Compatibility.IsWorksheets1Based = value;
                }
            }
        }

        /// <summary>
        /// Create a new file from a template
        /// </summary>
        /// <param name="template">An existing xlsx file to use as a template</param>
        /// <param name="password">The password to decrypt the package.</param>
        /// <returns></returns>
        private void CreateFromTemplate(FileInfo template, string password)
        {
            if (template != null) template.Refresh();
            if (template.Exists)
            {
                if(template.Length==0)
                {
                    throw new IOException($"{template.FullName} cannot be a zero-byte file.");
                }
                if (_stream==null)
                    _stream = RecyclableMemory.GetStream();

                var ms = RecyclableMemory.GetStream();
                if (password != null)
                {
                    Encryption.IsEncrypted = true;
                    Encryption.Password = password;
                    var encrHandler = new EncryptedPackageHandler();
                    ms.Dispose();
                    ms = encrHandler.DecryptPackage(template, Encryption);
                    encrHandler = null;
                }
                else
                {
                    WriteFileToStream(template.FullName, ms); 
                }
                try
                {
                    _zipPackage = new Packaging.ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    if (password == null && CompoundDocument.IsCompoundDocument(ms))
                    {
                        throw new Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
                    }
                    else
                    {
                        throw;
                    }
                }
                finally
                {
                    ms.Dispose();
				}
            }
            else
                throw new Exception("Passed invalid TemplatePath to Excel Template");
            //return newFile;
        }
        private void ConstructNewFile(string password)
        {
            var ms = RecyclableMemory.GetStream();
            if (_stream == null) _stream = RecyclableMemory.GetStream();
            if (File != null) File.Refresh();
            if (File != null && File.Exists && File.Length > 0)
            {
                if (password != null)
                {
                    var encrHandler = new EncryptedPackageHandler();
                    Encryption.IsEncrypted = true;
                    Encryption.Password = password;
                    ms.Dispose();
                    ms = encrHandler.DecryptPackage(File, Encryption);
                }
                else
                {
                    WriteFileToStream(File.FullName, ms);
                }
                try
                {
                    _zipPackage = new ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    if (password == null && CompoundDocument.IsCompoundDocument(File))
                    {
                        throw new Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
                    }
                    else
                    {
                        throw;
                    }
                }
    //            finally
    //            {
    //                ms.Dispose();
				//}
            }
            else
            {
                _zipPackage = new ZipPackage(ms);
                ms.Dispose();
                CreateBlankWb();
            }
        }
        /// <summary>
        /// Pull request from  perkuypers to read open Excel workbooks
        /// </summary>
        /// <param name="path">Path</param>
        /// <param name="stream">Stream</param>
        private static void WriteFileToStream(string path, Stream stream)
        {
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var buffer = new byte[4096];
                int read;
                while ((read = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    stream.Write(buffer, 0, read);
                }
            }
        }
        private void CreateBlankWb()
        {
            XmlDocument workbook = Workbook.WorkbookXml; // this will create the workbook xml in the package
            // create the relationship to the main part
            _zipPackage.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), Workbook.WorkbookUri), Packaging.TargetMode.Internal, schemaRelationships + "/officeDocument");
        }

        PictureStore _pictureStore = null;
        internal PictureStore PictureStore
        {
            get
            {
                if(_pictureStore==null)
                {
                    _pictureStore = new PictureStore(this);
                }
                return _pictureStore;
            }
        }
        /// <summary>
        /// Returns a reference to the package
        /// </summary>
        internal Packaging.ZipPackage ZipPackage { get { return (_zipPackage); } }
        ExcelEncryption _encryption=null;
        /// <summary>
        /// Information how and if the package is encrypted
        /// </summary>
        public ExcelEncryption Encryption
        {
            get
            {
                if (_encryption == null)
                {
                    _encryption = new ExcelEncryption();
                }
                return _encryption;
            }
        }
        private static LicenseContext? _licenseType = null;
        internal static bool _licenseSet = false;
        /// <summary>
        /// To use the EPPlus library in debug mode a Licensetype must be specified.
        /// Use LicenseContext.NonCommercial if you use EPPlus in an non commercial context.
        /// Use LicenseContext.Commercial if you have purchased an license to use EPPlus
        /// See https://epplussoftware.com/developers/licenseexception
        /// </summary>
        public static LicenseContext? LicenseContext
        {
            get
            {                
                return _licenseType;
            }
            set
            {
                _licenseType = value;
                _licenseSet = _licenseType != null;
            }
        }
        internal static bool IsLicenseSet(List<ExcelInitializationError> initErrors)
        {
            if(_licenseSet==true)
            {
                return true;
            }
            else
            {
                if (Debugger.IsAttached == false)   //This check is only performed if a debugger is attached. 
                {
                    _licenseSet = true;
                    return true;
                }
                var v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlusLicenseContext",EnvironmentVariableTarget.User, _configuration, initErrors);
                if(string.IsNullOrEmpty(v))
                {
                    v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlusLicenseContext", EnvironmentVariableTarget.Process, _configuration, initErrors);
                }
                bool inEnvironment;
                if (string.IsNullOrEmpty(v))
                {
#if (Core)
                    v = ExcelConfigurationReader.GetJsonConfigValue("EPPlus:ExcelPackage:LicenseContext", _configuration, initErrors);

#else
                    v = ExcelConfigurationReader.GetValueFromAppSettings("EPPlus:ExcelPackage:LicenseContext", _configuration, initErrors);
                    if(string.IsNullOrEmpty(v))
                    {
                        v = v = ExcelConfigurationReader.GetValueFromAppSettings("EPPlus:ExcelPackage.LicenseContext", _configuration, initErrors);
                    }
#endif
                    inEnvironment = false;
                }
                else
                {
                    inEnvironment = true;
                }

                if(string.IsNullOrEmpty(v))
                {
                    inEnvironment = false;
                    return false;
                }
                else
                {
                    v = v.Trim();
                    if (v.Equals("commercial", StringComparison.OrdinalIgnoreCase))
                    {
                        LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                        _licenseSet = true;
                        return _licenseSet;
                    }
                    else if (v.Equals("noncommercial", StringComparison.OrdinalIgnoreCase))
                    {
                        LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        _licenseSet = true;
                        return _licenseSet;
                    }
                }

                if(inEnvironment)
                    throw new LicenseException("LicenseContext is set to an invalid value in the environment variable 'EPPlusLicenseContext'. Please use Commercial or Noncommercial");
                else
                    throw new LicenseException("LicenseContext is set to an invalid value in the configuration file, Key: ExcelPackage.LicenseContext. Please use Commercial or Noncommercial");
            }
        }
		/// <summary>
		/// Returns a reference to the workbook component within the package.
		/// All worksheets and cells can be accessed through the workbook.
		/// </summary>
		public ExcelWorkbook Workbook
		{
			get
			{
                CheckNotDisposed();
                if (_workbook == null)
                {
                    if(IsLicenseSet(_initErrors)==false)
                    {
                        throw (new LicenseException("Please set the ExcelPackage.LicenseContext property. See https://epplussoftware.com/developers/licenseexception"));
                    }
                    var nsm = CreateDefaultNSM();

                    _workbook = new ExcelWorkbook(this, nsm);

                    _workbook.GetDefinedNames();
                    _workbook.LoadPivotTableCaches();

                }
                return (_workbook);
			}
		}

        /// <summary>
        /// Global configuration for the ExcelPackage class
        /// </summary>
        /// <param name="configHandler"></param>
        public static void Configure(Action<ExcelPackageConfiguration> configHandler)
        {
            configHandler(_configuration);
        }

        /// <summary>
        /// Errors that has been logged during initialization of the ExcelPackage class.
        /// </summary>
        public IEnumerable<ExcelInitializationError> InitializationErrors
        {
            get
            {
                return _initErrors;
            }
        }
        /// <summary>
        /// Automaticlly adjust drawing size when column width/row height are adjusted, depending on the drawings editBy property.
        /// Default True
        /// </summary>
        public bool DoAdjustDrawings
        {
            get;
            set;
        }
        private XmlNamespaceManager CreateDefaultNSM()
        {
            //  Create a NamespaceManager to handle the default namespace, 
            //  and create a prefix for the default namespace:
            NameTable nt = new NameTable();
            var ns = new XmlNamespaceManager(nt);
            ns.AddNamespace(string.Empty, ExcelPackage.schemaMain);
            ns.AddNamespace("d", schemaMain);            
            ns.AddNamespace("r", schemaRelationships);
            ns.AddNamespace("c", schemaChart);
            ns.AddNamespace("vt", schemaVt);
            // extended properties (app.xml)
            ns.AddNamespace("xp", schemaExtended);
            // custom properties
            ns.AddNamespace("ctp", schemaCustom);
            // core properties
            ns.AddNamespace("cp", schemaCore);
            // core property namespaces 
            ns.AddNamespace("dc", schemaDc);
            ns.AddNamespace("dcterms", schemaDcTerms);
            ns.AddNamespace("dcmitype", schemaDcmiType);
            ns.AddNamespace("xsi", schemaXsi);
            ns.AddNamespace("x14", schemaMainX14);
            ns.AddNamespace("x15", schemaMainX15);
            ns.AddNamespace("xm", schemaMainXm);
            ns.AddNamespace("xr", schemaXr);
            ns.AddNamespace("xr2", schemaXr2);
            ns.AddNamespace("mc", schemaMarkupCompatibility);
            ns.AddNamespace("tc", schemaThreadedComments);
            ns.AddNamespace("a14", schemaDrawings2010);
            ns.AddNamespace("xdr", schemaSheetDrawings);
            return ns;
        }
		
#region SavePart
		/// <summary>
		/// Saves the XmlDocument into the package at the specified Uri.
		/// </summary>
		/// <param name="uri">The Uri of the component</param>
		/// <param name="xmlDoc">The XmlDocument to save</param>
		internal void SavePart(Uri uri, XmlDocument xmlDoc)
		{
            Packaging.ZipPackagePart part = _zipPackage.GetPart(uri);
            var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            var xr = new XmlTextWriter(stream, Encoding.UTF8);
            xr.Formatting = Formatting.None;
            
            xmlDoc.Save(xr);
		}
        /// <summary>
		/// Saves the XmlDocument into the package at the specified Uri.
		/// </summary>
		/// <param name="uri">The Uri of the component</param>
		/// <param name="xmlDoc">The XmlDocument to save</param>
        internal void SaveWorkbook(Uri uri, XmlDocument xmlDoc)
		{
            Packaging.ZipPackagePart part = _zipPackage.GetPart(uri);
            if(Workbook.VbaProject==null)
            {
                if (part.ContentType != ContentTypes.contentTypeWorkbookDefault)
                {
                    part = _zipPackage.CreatePart(uri, ContentTypes.contentTypeWorkbookDefault, Compression);
                }
            }
            else
            {
                if (part.ContentType != ContentTypes.contentTypeWorkbookMacroEnabled)
                {
                    var rels = part.GetRelationships();
                    _zipPackage.DeletePart(uri);
                    part = ZipPackage.CreatePart(uri, ContentTypes.contentTypeWorkbookMacroEnabled);
                    foreach (ZipPackageRelationship rel in rels)
                    {
                        ZipPackage.DeleteRelationship(rel.Id);
                        part.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType);
                    }
                }
            }
            var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            var xr = new XmlTextWriter(stream, Encoding.UTF8);
            xr.Formatting = Formatting.None;

            xmlDoc.Save(xr);
		}

#endregion

#region Dispose
		/// <summary>
		/// Closes the package.
		/// </summary>
		public void Dispose()
		{
            if(_zipPackage != null)
            {
                if (_isExternalStream==false && _stream != null && (_stream.CanRead || _stream.CanWrite))
                {
                    CloseStream();
                }
                _zipPackage.Close();    
                if(_workbook != null)
                {
                    _workbook.Dispose();
                }
                _zipPackage = null;
                File = null;
                _workbook = null;
                _stream = null;
                _workbook = null;

                _isDisposed = true;
                if (Settings.DoGarbageCollectOnDispose)
                {
                    GC.Collect();
                }
            }
		}
        #endregion

        #region Save  // ExcelPackage save

        /// <summary>
        /// Saves all the components back into the package.
        /// This method recursively calls the Save method on all sub-components.
        /// We close the package after the save is done.
        /// </summary>
        public void Save()
        {
            CheckNotDisposed();
            try
            {
                if (_stream is MemoryStream && _stream.Length > 0)
                {
                    //Close any open memorystream and "renew" then. This can occure if the package is saved twice. 
                    //The stream is left open on save to enable the user to read the stream-property.
                    //Non-memorystream streams will leave the closing to the user before saving a second time.
                    CloseStream();
                }

                //Invoke before save delegates
                foreach (var action in BeforeSave)
                {
                    action.Invoke();
                }

                Workbook.Save();
                if (File == null)
                {
                    if(Encryption.IsEncrypted)
                    {
                        byte[] file;
                        using (var ms = RecyclableMemory.GetStream())
                        {
                            _zipPackage.Save(ms);
                            file = ms.ToArray();
                        }
                        EncryptedPackageHandler eph = new EncryptedPackageHandler();
                        using (var msEnc = eph.EncryptPackage(file, Encryption))
                        {
                            CopyStream(msEnc, ref _stream);
                        }
                    }   
                    else
                    {
                        _zipPackage.Save(_stream);
                    }
                    _stream.Flush();
                    _zipPackage.Close();
                }
                else
                {
                    if (System.IO.File.Exists(File.FullName))
                    {
                        try
                        {
                            System.IO.File.Delete(File.FullName);
                        }
                        catch (Exception ex)
                        {
                            throw (new Exception(string.Format("Error overwriting file {0}", File.FullName), ex));
                        }
                    }

                    _zipPackage.Save(_stream);
                    _zipPackage.Close();
                    if (Stream is MemoryStream)
                    {
                        using (var fi = new FileStream(File.FullName, FileMode.Create))
                        {
                            //EncryptPackage
                            if (Encryption.IsEncrypted)
                            {
                                byte[] file = ((MemoryStream)Stream).ToArray();
                                EncryptedPackageHandler eph = new EncryptedPackageHandler();

                                using (var ms = eph.EncryptPackage(file, Encryption))
                                {
                                    fi.Write(ms.ToArray(), 0, (int)ms.Length);
                                }
                            }
                            else
                            {
                                fi.Write(((MemoryStream)Stream).ToArray(), 0, (int)Stream.Length);
                            }
                            fi.Close();
                        }
                    }
                    else
                    {
                        System.IO.File.WriteAllBytes(File.FullName, GetAsByteArray(false));
                    }
                }
            }
            catch (Exception ex)
            {
                if (File == null)
                {
                    throw;
                }
                else
                {
                        throw (new InvalidOperationException(string.Format("Error saving file {0}", File.FullName), ex));
                }
            }
        }
        /// <summary>
        /// Saves all the components back into the package.
        /// This method recursively calls the Save method on all sub-components.
        /// The package is closed after it has ben saved
        /// Supply a password to encrypt the workbook with. 
        /// </summary>
        /// <param name="password">This parameter overrides the Workbook.Encryption.Password.</param>
        public void Save(string password)
		{
            Encryption.Password = password;
            Save();
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved        
        /// </summary>
        /// <param name="file">The file location</param>
        public void SaveAs(FileInfo file)
        {
            File = file;
            Save();
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved        
        /// </summary>
        /// <param name="filePath">The file location</param>
        public void SaveAs(string filePath)
        {
            SaveAs(new FileInfo(filePath));
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="file">The file</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        public void SaveAs(FileInfo file, string password)
        {
            File = file;
            Encryption.Password = password;
            Save();
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="filePath">The file</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        public void SaveAs(string filePath, string password)
        {
            SaveAs(new FileInfo(filePath), password);
        }
        /// <summary>
        /// Copies the Package to the Outstream
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="OutputStream">The stream to copy the package to</param>
        public void SaveAs(Stream OutputStream)
        {
            File = null;
            Save();

            if (OutputStream != _stream)
            {
                CopyStream(_stream, ref OutputStream);
            }
        }
        /// <summary>
        /// Copies the Package to the Outstream
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="OutputStream">The stream to copy the package to</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        public void SaveAs(Stream OutputStream, string password)
        {
            Encryption.Password = password;
            SaveAs(OutputStream);
        }

        /// <summary>
        /// The output file. Null if no file is used
        /// </summary>
        public FileInfo File { get; set; } = null;
        /// <summary>
        /// Close the internal stream
        /// </summary>
        internal void CloseStream()
        {
            // Issue15252: Clear output buffer
            if (_stream != null)
            {
                _stream.Close();
                _stream.Dispose();
            }

            _stream = RecyclableMemory.GetStream();
        }
        /// <summary>
        /// The output stream. This stream is the not the encrypted package.
        /// To get the encrypted package use the SaveAs(stream) method.
        /// </summary>
        public Stream Stream
        {
            get
            {
                return _stream;
            }
        }
#endregion
        /// <summary>
        /// Compression option for the package
        /// </summary>        
        public CompressionLevel Compression 
        { 
            get
            {
                return ZipPackage.Compression;
            }
            set
            {
                ZipPackage.Compression = value;
            }
        }
        CompatibilitySettings _compatibility = null;
        /// <summary>
        /// Compatibility settings for older versions of EPPlus.
        /// </summary>
        public CompatibilitySettings Compatibility
        {
            get
            {
                if(_compatibility==null)
                {
                    _compatibility=new CompatibilitySettings(this);
                }
                return _compatibility;
            }
        }
        /// <summary>
        /// Package generic settings
        /// </summary>
        public ExcelPackageSettings Settings
        {
            get;
            private set;
        } = new ExcelPackageSettings();
        #region GetXmlFromUri
        /// <summary>
        /// Get the XmlDocument from an URI
        /// </summary>
        /// <param name="uri">The Uri to the part</param>
        /// <returns>The XmlDocument</returns>
        internal XmlDocument GetXmlFromUri(Uri uri)
		{
			XmlDocument xml = new XmlDocument();
			Packaging.ZipPackagePart part = _zipPackage.GetPart(uri);
            XmlHelper.LoadXmlSafe(xml, part.GetStream()); 
			return (xml);
		}
        #endregion
        #region GetAsByteArray
        /// <summary>
        /// Saves and returns the Excel files as a bytearray.
        /// Note that the package is closed upon save
        /// </summary>
        /// <example>      
        /// Example how to return a document from a Webserver...
        /// <code> 
        ///  ExcelPackage package=new ExcelPackage();
        ///  /**** ... Create the document ****/
        ///  Byte[] bin = package.GetAsByteArray();
        ///  Response.ContentType = "Application/vnd.ms-Excel";
        ///  Response.AddHeader("content-disposition", "attachment;  filename=TheFile.xlsx");
        ///  Response.BinaryWrite(bin);
        /// </code>
        /// </example>
        /// <returns></returns>
        public byte[] GetAsByteArray()
        {
           return GetAsByteArray(true);
        }
        /// <summary>
        /// Saves and returns the Excel files as a bytearray
        /// Note that the package is closed upon save
        /// </summary>
        /// <example>      
        /// Example how to return a document from a Webserver...
        /// <code> 
        ///  ExcelPackage package=new ExcelPackage();
        ///  /**** ... Create the document ****/
        ///  Byte[] bin = package.GetAsByteArray();
        ///  Response.ContentType = "Application/vnd.ms-Excel";
        ///  Response.AddHeader("content-disposition", "attachment;  filename=TheFile.xlsx");
        ///  Response.BinaryWrite(bin);
        /// </code>
        /// </example>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <returns></returns>
        public byte[] GetAsByteArray(string password)
        {
            if (password != null)
            {
                Encryption.Password = password;
            }
            return GetAsByteArray(true);
        }
        internal byte[] GetAsByteArray(bool save)
        {
            CheckNotDisposed();
            if (save)
            {
                Workbook.Save();
                _zipPackage.Close();
                if(_stream is MemoryStream && _stream.Length>0)
                {
                    CloseStream();
                }
                _zipPackage.Save(_stream);
            }
            Byte[] byRet = new byte[Stream.Length];
            long pos = Stream.Position;            
            Stream.Seek(0, SeekOrigin.Begin);
            Stream.Read(byRet, 0, (int)Stream.Length);

            //Encrypt Workbook?
            if (Encryption.IsEncrypted)
            {
                EncryptedPackageHandler eph=new EncryptedPackageHandler();
                using (var ms = eph.EncryptPackage(byRet, Encryption))
                {
                    byRet = ms.ToArray();
                }
            }

            Stream.Seek(pos, SeekOrigin.Begin);
            Stream.Close();
            return byRet;
        }

        private void CheckNotDisposed()
        {
            if(_isDisposed)
            {
                throw (new ObjectDisposedException("ExcelPackage", "Package has been disposed"));
            }
        }
        #endregion
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        public void Load(Stream input)
        {
            Load(input, RecyclableMemory.GetStream(), null);            
        }
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="Password">The password to decrypt the document</param>
        public void Load(Stream input, string Password)
        {
            Load(input, RecyclableMemory.GetStream(), Password);
        }   
        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>    
        /// <param name="output"></param>
        /// <param name="Password"></param>
        private void Load(Stream input, Stream output, string Password)
        {
            ReleaseResources();
            if (input.CanSeek && input.Length == 0) // Template is blank, Construct new
            {
                _stream = output;
                ConstructNewFile(Password);
            }
            else
            {
                Stream ms;
                this._stream = output;
                if (Password != null)
                {
                    Stream encrStream = RecyclableMemory.GetStream();
                    CopyStream(input, ref encrStream);
                    EncryptedPackageHandler eph = new EncryptedPackageHandler();
                    Encryption.Password = Password;
                    ms = eph.DecryptPackage((MemoryStream)encrStream, Encryption);
                    encrStream.Dispose();
                }
                else
                {
                    ms = RecyclableMemory.GetStream();
                    CopyStream(input, ref ms);
                }

                try
                {
                    //this._package = Package.Open(this._stream, FileMode.Open, FileAccess.ReadWrite);
                    _zipPackage = new Packaging.ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    EncryptedPackageHandler eph = new EncryptedPackageHandler();
                    if (Password == null && CompoundDocument.IsCompoundDocument((MemoryStream)_stream))
                    {
                        throw new Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
                    }
                    else
                    {
                        throw;
                    }
                }
                finally
                {
                    ms.Dispose();
				}
            }            
            //Clear the workbook so that it gets reinitialized next time
            this._workbook = null;
        }

        private void ReleaseResources()
        {
            //Release some resources:
            if (_zipPackage != null)
            {
                _zipPackage.Close();
                _zipPackage = null;
            }
            if (_stream != null)
            {
                _stream.Close();
                _stream.Dispose();
                _stream = null;
            }
            _isExternalStream = true;
            _isDisposed = false;
        }

        static object _lock=new object();
        internal int _worksheetAdd=0;
        /// <summary>
        /// Copies the input stream to the output stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="outputStream">The output stream.</param>
        internal static void CopyStream(Stream inputStream, ref Stream outputStream)
        {
            if (!inputStream.CanRead)
            {
                throw (new Exception("Cannot read from the input stream"));
            }
            if (!outputStream.CanWrite)
            {
                throw (new Exception("Cannot write to the output stream"));
            }
            if (inputStream.CanSeek)
            {
                inputStream.Seek(0, SeekOrigin.Begin);
            }

                const int bufferLength = 8096;
                var buffer = new Byte[bufferLength];
                lock (_lock)
                {
                    int bytesRead = inputStream.Read(buffer, 0, bufferLength);
                    // write the required bytes
                    while (bytesRead > 0)
                    {
                        outputStream.Write(buffer, 0, bytesRead);
                        bytesRead = inputStream.Read(buffer, 0, bufferLength);
                    }
                    outputStream.Flush();
            }
        }
    }
}