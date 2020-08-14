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
using System.Text;
using System.Globalization;
using OfficeOpenXml.VBA;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Table;
using System.Linq;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml
{
	#region Public Enum ExcelCalcMode
	/// <summary>
	/// How the application should calculate formulas in the workbook
	/// </summary>
	public enum ExcelCalcMode
	{
		/// <summary>
		/// Indicates that calculations in the workbook are performed automatically when cell values change. 
		/// The application recalculates those cells that are dependent on other cells that contain changed values. 
		/// This mode of calculation helps to avoid unnecessary calculations.
		/// </summary>
		Automatic,
		/// <summary>
		/// Indicates tables be excluded during automatic calculation
		/// </summary>
		AutomaticNoTable,
		/// <summary>
		/// Indicates that calculations in the workbook be triggered manually by the user. 
		/// </summary>
		Manual
	}
	#endregion

	/// <summary>
	/// Represents the Excel workbook and provides access to all the 
	/// document properties and worksheets within the workbook.
	/// </summary>
	public sealed class ExcelWorkbook : XmlHelper, IDisposable
	{
		internal class SharedStringItem
		{
			internal int pos;
			internal string Text;
			internal bool isRichText = false;
		}
		#region Private Properties
		internal ExcelPackage _package;
		internal ExcelWorksheets _worksheets;
		private OfficeProperties _properties;

		private ExcelStyles _styles;
		internal HashSet<string> _tableSlicerNames = new HashSet<string>();
		internal HashSet<string> _pivotTableSlicerNames = new HashSet<string>();

        internal string GetTableSlicerName(string name)
        {
            return GetUniqueName(name, _tableSlicerNames);
        }

        internal bool GetPivotCacheFromAddress(string fullAddress, out PivotTableCacheInternal cacheReference)
        {
			if (_package.Compatibility.SharePivotTableCacheForSameRange)
			{
				if(_pivotTableCaches.TryGetValue(fullAddress, out PivotTableCacheRangeInfo cacheInfo))
				{
					cacheReference=cacheInfo.PivotCaches[0];
					return true;
				}
			}
			cacheReference = null;
			return false;

		}

		internal string GetPivotTableSlicerName(string name)
		{
			return GetUniqueName(name, _pivotTableSlicerNames);
		}

		private string GetUniqueName(string name, HashSet<string> hs)
        {
            var n = name;
            var ix = 1;
            while (hs.Contains(n))
            {
                n = name + $"{ix}";
            }
            return n;
        }
        #endregion

        #region ExcelWorkbook Constructor
        /// <summary>
        /// Creates a new instance of the ExcelWorkbook class.
        /// </summary>
        /// <param name="package">The parent package</param>
        /// <param name="namespaceManager">NamespaceManager</param>
        internal ExcelWorkbook(ExcelPackage package, XmlNamespaceManager namespaceManager) :
			base(namespaceManager)
		{
			_package = package;
			SetUris();

			_names = new ExcelNamedRangeCollection(this);
			_namespaceManager = namespaceManager;
			TopNode = WorkbookXml.DocumentElement;
			SchemaNodeOrder = new string[] { "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "functionPrototypes", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst" };
		    FullCalcOnLoad = true;  //Full calculation on load by default, for both new workbooks and templates.
			GetSharedStrings();
		}

		private void SetUris()
		{
			foreach(var rel in _package.Package.GetRelationships())
			{
				if (rel.RelationshipType == ExcelPackage.schemaRelationships+ "/officeDocument")
				{
					WorkbookUri = rel.TargetUri;
					break;
				}
			}

			if(WorkbookUri==null)
			{
				WorkbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
			}
			else
			{
				foreach (var rel in Part.GetRelationships())
				{
					if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/sharedStrings")
					{
						SharedStringsUri = UriHelper.ResolvePartUri(WorkbookUri, rel.TargetUri);
					}
					else if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/styles")
					{
						StylesUri = UriHelper.ResolvePartUri(WorkbookUri, rel.TargetUri);
					}
				}
			}

			if(SharedStringsUri==null)
				SharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
			if (StylesUri == null)
				StylesUri = new Uri("/xl/styles.xml", UriKind.Relative);

		}
		#endregion

		internal Dictionary<string, SharedStringItem> _sharedStrings = new Dictionary<string, SharedStringItem>(); //Used when reading cells.
		internal List<SharedStringItem> _sharedStringsList = new List<SharedStringItem>(); //Used when reading cells.
		internal ExcelNamedRangeCollection _names;
		internal int _nextDrawingID = 2;
		internal int _nextTableID = int.MinValue;
		internal int _nextPivotCacheId=1;
        internal int GetNewPivotCacheId()
        {
			return _nextPivotCacheId++;
		}
		internal void SetNewPivotCacheId(int value)
        {
			if(value >= _nextPivotCacheId) _nextPivotCacheId = value+1;
        }
		internal int _nextPivotTableID = int.MinValue;
		internal XmlNamespaceManager _namespaceManager;
        internal FormulaParser _formulaParser = null;
	    internal FormulaParserManager _parserManager;
        internal CellStore<List<Token>> _formulaTokens;
		internal class PivotTableCacheRangeInfo
        {
            public string Address { get; set; }
            public List<PivotTableCacheInternal> PivotCaches { get; set; }
        }
		internal Dictionary<string, PivotTableCacheRangeInfo> _pivotTableCaches = new Dictionary<string, PivotTableCacheRangeInfo>();
		/// <summary>
		/// Read shared strings to list
		/// </summary>
		private void GetSharedStrings()
		{
			if (_package.Package.PartExists(SharedStringsUri))
			{
				var xml = _package.GetXmlFromUri(SharedStringsUri);
				XmlNodeList nl = xml.SelectNodes("//d:sst/d:si", NameSpaceManager);
				_sharedStringsList = new List<SharedStringItem>();
				if (nl != null)
				{
					foreach (XmlNode node in nl)
					{
						XmlNode n = node.SelectSingleNode("d:t", NameSpaceManager);
						if (n != null)
						{
                            _sharedStringsList.Add(new SharedStringItem() { Text = ConvertUtil.ExcelDecodeString(n.InnerText) });
						}
						else
						{
							_sharedStringsList.Add(new SharedStringItem() { Text = node.InnerXml, isRichText = true });
						}
					}
				}
                //Delete the shared string part, it will be recreated when the package is saved.
                foreach (var rel in Part.GetRelationships())
                {
                    if (rel.TargetUri.OriginalString.EndsWith("sharedstrings.xml", StringComparison.OrdinalIgnoreCase))
                    {
                        Part.DeleteRelationship(rel.Id);
                        break;
                    }
                }                
                _package.Package.DeletePart(SharedStringsUri); //Remove the part, it is recreated when saved.
			}
		}
		internal void GetDefinedNames()
		{
			XmlNodeList nl = WorkbookXml.SelectNodes("//d:definedNames/d:definedName", NameSpaceManager);
			if (nl != null)
			{
				foreach (XmlElement elem in nl)
				{ 
					string fullAddress = elem.InnerText;

					int localSheetID;
					ExcelWorksheet nameWorksheet;
					
                    if(!int.TryParse(elem.GetAttribute("localSheetId"), NumberStyles.Number, CultureInfo.InvariantCulture, out localSheetID))
					{
						localSheetID = -1;
						nameWorksheet=null;
					}
					else
					{
						nameWorksheet=Worksheets[localSheetID + _package._worksheetAdd];
					}

					var addressType = ExcelAddressBase.IsValid(fullAddress);
					ExcelRangeBase range;
					ExcelNamedRange namedRange;

					if (fullAddress.IndexOf("[") == 0)
					{
						int start = fullAddress.IndexOf("[");
						int end = fullAddress.IndexOf("]", start);
						if (start >= 0 && end >= 0)
						{

							string externalIndex = fullAddress.Substring(start + 1, end - start - 1);
							int index;
							if (int.TryParse(externalIndex, NumberStyles.Any, CultureInfo.InvariantCulture, out index))
							{
								if (index > 0 && index <= _externalReferences.Count)
								{
									fullAddress = fullAddress.Substring(0, start) + "[" + _externalReferences[index - 1] + "]" + fullAddress.Substring(end + 1);
								}
							}
						}
					}

					if (addressType == ExcelAddressBase.AddressType.Invalid || addressType == ExcelAddressBase.AddressType.InternalName || addressType == ExcelAddressBase.AddressType.ExternalName || addressType==ExcelAddressBase.AddressType.Formula || addressType==ExcelAddressBase.AddressType.ExternalAddress)    //A value or a formula
					{
						double value;
						range = new ExcelRangeBase(this, nameWorksheet, elem.GetAttribute("name"), true);
						if (nameWorksheet == null)
						{
							namedRange = _names.Add(elem.GetAttribute("name"), range);
						}
						else
						{
							namedRange = nameWorksheet.Names.Add(elem.GetAttribute("name"), range);
						}
						
						if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "\"")) //String value
						{
							namedRange.NameValue = fullAddress.Substring(1,fullAddress.Length-2);
						}
						else if (double.TryParse(fullAddress, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
						{
							namedRange.NameValue = value;
						}
						else
						{
                            //if (addressType == ExcelAddressBase.AddressType.ExternalAddress || addressType == ExcelAddressBase.AddressType.ExternalName)
                            //{
                            //    var r = new ExcelAddress(fullAddress);
                            //    namedRange.NameFormula = '\'[' + r._wb
                            //}
                            //else
                            //{
                                namedRange.NameFormula = fullAddress;
                            //}
						}
					}
					else
					{
						ExcelAddress addr = new ExcelAddress(fullAddress, _package, null);
						if (localSheetID > -1)
						{
							if (string.IsNullOrEmpty(addr._ws))
							{
								namedRange = Worksheets[localSheetID + _package._worksheetAdd].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[localSheetID + _package._worksheetAdd], fullAddress, false));
							}
							else
							{
								namedRange = Worksheets[localSheetID + _package._worksheetAdd].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[addr._ws], fullAddress, false));
							}
						}
						else
						{
							var ws = Worksheets[addr._ws];
							namedRange = _names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, ws, fullAddress, false));
						}
					}
					if (elem.GetAttribute("hidden") == "1" && namedRange != null) namedRange.IsNameHidden = true;
					if(!string.IsNullOrEmpty(elem.GetAttribute("comment"))) namedRange.NameComment=elem.GetAttribute("comment");
				}
			}
		}

        internal ExcelRangeBase GetRange(ExcelWorksheet ws, string function)
        {
            switch (ExcelAddressBase.IsValid(function))
            {
                case ExcelAddressBase.AddressType.InternalAddress:
                    var addr = new ExcelAddress(function);
                    if (string.IsNullOrEmpty(addr.WorkSheetName))
                    {
                        return ws.Cells[function];
                    }
                    else
                    {
                        var otherWs = Worksheets[addr.WorkSheetName];
                        if (otherWs == null)
                        {
                            return null;
                        }
                        else
                        {
                            return otherWs.Cells[addr.Address];
                        }
                    }
                case ExcelAddressBase.AddressType.InternalName:
                    if (Names.ContainsKey(function))
                    {
                        return Names[function];
                    }
                    else if(ws.Names.ContainsKey(function))
                    {
                        return ws.Names[function];
                    }
                    else if(ws.Tables[function]!=null)
                    {                        
                        return ws.Cells[ws.Tables[function].Address.Address];
                    }
                    else
                    {
                        var nameAddr = new ExcelAddress(function);
                        if (string.IsNullOrEmpty(nameAddr.WorkSheetName))
                        {
                            return null;
                        }
                        else
                        {
                            var otherWs = Worksheets[nameAddr.WorkSheetName];
                            if (otherWs != null && otherWs.Names.ContainsKey(nameAddr.Address))
                            {
                                return otherWs.Names[nameAddr.Address];
                            }
                            return null;
                        }
                    }
                case ExcelAddressBase.AddressType.Formula:
                    return null;
                default:
                    return null;
            }           
        }
        #region Worksheets
        /// <summary>
        /// Provides access to all the worksheets in the workbook.
        /// Note: Worksheets index either starts by 0 or 1 depending on the Excelpackage.Compatibility.IsWorksheets1Based property.
        /// Default is 1 for .Net 3.5 and .Net 4 and 0 for .Net Core.
        /// </summary>
        public ExcelWorksheets Worksheets
		{
			get
			{
				if (_worksheets == null)
				{
					var sheetsNode = _workbookXml.DocumentElement.SelectSingleNode("d:sheets", _namespaceManager);
					if (sheetsNode == null)
					{
						sheetsNode = CreateNode("d:sheets");
					}
					
					_worksheets = new ExcelWorksheets(_package, _namespaceManager, sheetsNode);
				}
				return (_worksheets);
			}
		}
		#endregion

		/// <summary>
		/// Provides access to named ranges
		/// </summary>
		public ExcelNamedRangeCollection Names
		{
			get
			{
				return _names;
			}
		}
		#region Workbook Properties
		decimal _standardFontWidth = decimal.MinValue;
        string _fontID = "";
        internal FormulaParser FormulaParser
        {
            get
            {
                if (_formulaParser == null)
                {
                    _formulaParser = new FormulaParser(new EpplusExcelDataProvider(_package));
                }
                return _formulaParser;
            }
        }
        /// <summary>
        /// Manage the formula parser.
        /// Add your own functions or replace native ones, parse formulas or attach a logger.
        /// </summary>
	    public FormulaParserManager FormulaParserManager
	    {
	        get
	        {
	            if (_parserManager == null)
	            {
	                _parserManager = new FormulaParserManager(FormulaParser);
	            }
	            return _parserManager;
	        }
	    }
        /// <summary>
		/// Max font width for the workbook
        /// <remarks>This method uses GDI. If you use Azure or another environment that does not support GDI, you have to set this value manually if you don't use the standard Calibri font</remarks>
		/// </summary>
		public decimal MaxFontWidth 
		{
			get
			{
                var ix = Styles.NamedStyles.FindIndexByID("Normal");
                if (ix >= 0)
                {
					var font = Styles.NamedStyles[ix].Style.Font;
					if (font.Index == int.MinValue) font.Index=0;
					if (_standardFontWidth == decimal.MinValue || _fontID != font.Id)
				    {
                        try
                        {
                            _standardFontWidth = GetWidthPixels(font.Name, font.Size);
                            _fontID = Styles.NamedStyles[ix].Style.Font.Id;
                        }
                        catch   //Error, Font missing and Calibri removed in dictionary
                        {
                            _standardFontWidth = (int)(font.Size * (2D / 3D)); //Aprox for Calibri.
                        }
                    }
                }
                else
                {
                    _standardFontWidth = 7; //Calibri 11
                }
                return _standardFontWidth;
			}
            set
            {
                _standardFontWidth = value;
            }
		}

	    internal static decimal GetWidthPixels(string fontName, float fontSize)
	    {
            Dictionary<float, FontSizeInfo> font;
            if (FontSize.FontHeights.ContainsKey(fontName))
            {
                font = FontSize.FontHeights[fontName];
            }
            else
            {
                font = FontSize.FontHeights["Calibri"];
            }

            if (font.ContainsKey(fontSize))
            {
                return Convert.ToDecimal(font[fontSize].Width);
            }
            else
            {
                float min = -1, max = 500;
                foreach (var size in font)
                {
                    if (min < size.Key && size.Key < fontSize)
                    {
                        min = size.Key;
                    }
                    if (max > size.Key && size.Key > fontSize)
                    {
                        max = size.Key;
                    }
                }
                if (min == max)
                {
                    return Convert.ToDecimal(font[min].Width);
                }
                else
                {
                    return Convert.ToDecimal(font[min].Height + (font[max].Height - font[min].Height) * ((fontSize - min) / (max - min)));
                }
            }
        }

	    ExcelProtection _protection = null;
		/// <summary>
		/// Access properties to protect or unprotect a workbook
		/// </summary>
		public ExcelProtection Protection
		{
			get
			{
				if (_protection == null)
				{
					_protection = new ExcelProtection(NameSpaceManager, TopNode, this);
					_protection.SchemaNodeOrder = SchemaNodeOrder;
				}
				return _protection;
			}
		}        
		ExcelWorkbookView _view = null;
		/// <summary>
		/// Access to workbook view properties
		/// </summary>
		public ExcelWorkbookView View
		{
			get
			{
				if (_view == null)
				{
					_view = new ExcelWorkbookView(NameSpaceManager, TopNode, this);
				}
				return _view;
			}
		}
        ExcelVbaProject _vba = null;
        /// <summary>
        /// A reference to the VBA project.
        /// Null if no project exists.
        /// Use Workbook.CreateVBAProject to create a new VBA-Project
        /// </summary>
        public ExcelVbaProject VbaProject
        {
            get
            {
                if (_vba == null)
                {
                    if(_package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
                    {
                        _vba = new ExcelVbaProject(this);
                    }
                }
                return _vba;
            }
        }
		/// <summary>
		/// Remove the from the file VBA project.
		/// </summary>
		public void RemoveVBAProject()
		{
			if (_vba != null)
			{
				_vba.RemoveMe();
				_vba = null;
			}
		}

		/// <summary>
		/// Create an empty VBA project.
		/// </summary>
		public void CreateVBAProject()
        {
            if (_vba != null || _package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
            {
                throw (new InvalidOperationException("VBA project already exists."));
            }
                        
            _vba = new ExcelVbaProject(this);
            _vba.Create();
				}
		/// <summary>
		/// URI to the workbook inside the package
		/// </summary>
		internal Uri WorkbookUri { get; private set; }
		/// <summary>
        /// URI to the styles inside the package
		/// </summary>
		internal Uri StylesUri { get; private set; }
		/// <summary>
        /// URI to the shared strings inside the package
		/// </summary>
		internal Uri SharedStringsUri { get; private set; }
		/// <summary>
		/// Returns a reference to the workbook's part within the package
		/// </summary>
		internal Packaging.ZipPackagePart Part { get { return (_package.Package.GetPart(WorkbookUri)); } }
		
		#region WorkbookXml
		private XmlDocument _workbookXml;
		/// <summary>
		/// Provides access to the XML data representing the workbook in the package.
		/// </summary>
		public XmlDocument WorkbookXml
		{
			get
			{
				if (_workbookXml == null)
				{
					CreateWorkbookXml(_namespaceManager);
				}
				return (_workbookXml);
			}
		}
        const string codeModuleNamePath = "d:workbookPr/@codeName";
        internal string CodeModuleName
        {
            get
            {
                return GetXmlNodeString(codeModuleNamePath);
            }
            set
            {
                SetXmlNodeString(codeModuleNamePath,value);
            }
        }
        internal void CodeNameChange(string value)
        {
            CodeModuleName = value;
        }
        /// <summary>
        /// The VBA code module if the package has a VBA project. Otherwise this propery is null.
        /// <seealso cref="CreateVBAProject"/>
        /// </summary>
        public VBA.ExcelVBAModule CodeModule
        {
            get
            {
                if (VbaProject != null)
                {
                    return VbaProject.Modules[CodeModuleName];
                }
                else
                {
                    return null;
                }
            }
        }

        const string date1904Path = "d:workbookPr/@date1904";
        internal const double date1904Offset = 365.5 * 4;  // offset to fix 1900 and 1904 differences, 4 OLE years
        private bool? date1904Cache = null;

        internal bool ExistsPivotCache(int cacheID, ref int newID)
        {
            newID = cacheID;
            var ret = true;
            foreach (var ws in Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                foreach (var pt in ws.PivotTables)
                {
                    if(pt.CacheId==cacheID)
                    {
                        ret=false;
                    }
                    if(pt.CacheId>=newID)
                    {
                        newID = pt.CacheId+1;
                    }
                }
            }
            if (ret) newID = cacheID;   //Not Found, return same ID
            return ret;
        }

        /// <summary>
        /// The date systems used by Microsoft Excel can be based on one of two different dates. By default, a serial number of 1 in Microsoft Excel represents January 1, 1900.
        /// The default for the serial number 1 can be changed to represent January 2, 1904.
        /// This option was included in Microsoft Excel for Windows to make it compatible with Excel for the Macintosh, which defaults to January 2, 1904.
        /// </summary>
        public bool Date1904
        {
            get
            {
                if (date1904Cache == null)
                {
                    date1904Cache = GetXmlNodeBool(date1904Path, false);
                }
                return date1904Cache.Value;
            }
            set
            {
                if (Date1904 != value)
                {
                    // Like Excel when the option it's changed update it all cells with Date format
                    foreach (var ws in Worksheets)
                    {
                        if (ws is ExcelChartsheet) continue;
                        ws.UpdateCellsWithDate1904Setting();
                    }
                }
                date1904Cache = value;
                SetXmlNodeBool(date1904Path, value, false);
            }
        }
     
       
		/// <summary>
		/// Create or read the XML for the workbook.
		/// </summary>
		private void CreateWorkbookXml(XmlNamespaceManager namespaceManager)
		{
			if (_package.Package.PartExists(WorkbookUri))
				_workbookXml = _package.GetXmlFromUri(WorkbookUri);
			else
			{
				// create a new workbook part and add to the package
				Packaging.ZipPackagePart partWorkbook = _package.Package.CreatePart(WorkbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", _package.Compression);

				// create the workbook
				_workbookXml = new XmlDocument(namespaceManager.NameTable);                
				
				_workbookXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
				// create the workbook element
				XmlElement wbElem = _workbookXml.CreateElement("workbook", ExcelPackage.schemaMain);

				// Add the relationships namespace
				wbElem.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);

				_workbookXml.AppendChild(wbElem);

				// create the bookViews and workbooks element
				XmlElement bookViews = _workbookXml.CreateElement("bookViews", ExcelPackage.schemaMain);
				wbElem.AppendChild(bookViews);
				XmlElement workbookView = _workbookXml.CreateElement("workbookView", ExcelPackage.schemaMain);
				bookViews.AppendChild(workbookView);

				// save it to the package
				StreamWriter stream = new StreamWriter(partWorkbook.GetStream(FileMode.Create, FileAccess.Write));
				_workbookXml.Save(stream);
				//stream.Close();
				_package.Package.Flush();
			}
		}
		#endregion
		#region StylesXml
		private XmlDocument _stylesXml;
		/// <summary>
		/// Provides access to the XML data representing the styles in the package. 
		/// </summary>
		public XmlDocument StylesXml
		{
			get
			{
				if (_stylesXml == null)
				{
					if (_package.Package.PartExists(StylesUri))
						_stylesXml = _package.GetXmlFromUri(StylesUri);
					else
					{
						// create a new styles part and add to the package
						Packaging.ZipPackagePart part = _package.Package.CreatePart(StylesUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", _package.Compression);
						// create the style sheet

						StringBuilder xml = new StringBuilder("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
						xml.Append("<numFmts />");
						xml.Append("<fonts count=\"1\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts>");
						xml.Append("<fills><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills>");
						xml.Append("<borders><border><left /><right /><top /><bottom /><diagonal /></border></borders>");
						xml.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs>");
						xml.Append("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\" /></cellXfs>");
						xml.Append("<cellStyles><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles>");
                        xml.Append("<dxfs count=\"0\" />");
                        xml.Append("</styleSheet>");
						
						_stylesXml = new XmlDocument();
						_stylesXml.LoadXml(xml.ToString());
						
						//Save it to the package
						StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));

						_stylesXml.Save(stream);
						//stream.Close();
						_package.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, StylesUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/styles");
						_package.Package.Flush();
					}
				}
				return (_stylesXml);
			}
			set
			{
				_stylesXml = value;
			}
		}
		/// <summary>
		/// Package styles collection. Used internally to access style data.
		/// </summary>
		public ExcelStyles Styles
		{
			get
			{
				if (_styles == null)
				{
					_styles = new ExcelStyles(NameSpaceManager, StylesXml, this);
				}
				return _styles;
			}
		}
		#endregion

		#region Office Document Properties
		/// <summary>
		/// The office document properties
		/// </summary>
		public OfficeProperties Properties
		{
			get 
			{
				if (_properties == null)
				{
					//  Create a NamespaceManager to handle the default namespace, 
					//  and create a prefix for the default namespace:                   
					_properties = new OfficeProperties(_package, NameSpaceManager);
				}
				return _properties;
			}
		}
		#endregion

		#region CalcMode
		private string CALC_MODE_PATH = "d:calcPr/@calcMode";
		/// <summary>
		/// Calculation mode for the workbook.
		/// </summary>
		public ExcelCalcMode CalcMode
		{
			get
			{
				string calcMode = GetXmlNodeString(CALC_MODE_PATH);
				switch (calcMode)
				{
					case "autoNoTable":
						return ExcelCalcMode.AutomaticNoTable;
					case "manual":
						return ExcelCalcMode.Manual;
					default:
						return ExcelCalcMode.Automatic;
                        
				}
			}
			set
			{
				switch (value)
				{
					case ExcelCalcMode.AutomaticNoTable:
						SetXmlNodeString(CALC_MODE_PATH, "autoNoTable") ;
						break;
					case ExcelCalcMode.Manual:
						SetXmlNodeString(CALC_MODE_PATH, "manual");
						break;
					default:
						SetXmlNodeString(CALC_MODE_PATH, "auto");
						break;

				}
			}
			#endregion
		}

        private const string FULL_CALC_ON_LOAD_PATH = "d:calcPr/@fullCalcOnLoad";
        /// <summary>
        /// Should Excel do a full calculation after the workbook has been loaded?
        /// <remarks>This property is always true for both new workbooks and loaded templates(on load). If this is not the wanted behavior set this property to false.</remarks>
        /// </summary>
        public bool FullCalcOnLoad
	    {
	        get
	        {
                return GetXmlNodeBool(FULL_CALC_ON_LOAD_PATH);   
	        }
	        set
	        {
                SetXmlNodeBool(FULL_CALC_ON_LOAD_PATH, value);
	        }
	    }

        ExcelThemeManager _theme = null;
        /// <summary>
        /// Create and manage the theme for the workbook.
        /// </summary>
        public ExcelThemeManager ThemeManager
        {
            get
            {
                if(_theme==null)
                {
                    _theme = new ExcelThemeManager(this);
                }
                return _theme;
            }
        }
        const string defaultThemeVersionPath = "d:workbookPr/@defaultThemeVersion";
        /// <summary>
        /// The default version of themes to apply in the workbook
        /// </summary>
        public int? DefaultThemeVersion
        {
            get
            {
                return GetXmlNodeIntNull(defaultThemeVersionPath);
            }
            set
            {
                if(value is null)
                {
                    DeleteNode(defaultThemeVersionPath);
                }
                else
                {
                    SetXmlNodeString(defaultThemeVersionPath, value.ToString());
                }
            }
        }
        #endregion
        #region Workbook Private Methods

        #region Save // Workbook Save
        /// <summary>
        /// Saves the workbook and all its components to the package.
        /// For internal use only!
        /// </summary>
        internal void Save()  // Workbook Save
		{
			if (Worksheets.Count == 0)
				throw new InvalidOperationException("The workbook must contain at least one worksheet");

			DeleteCalcChain();

            if (_vba == null && !_package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
            {
                if (Part.ContentType != ExcelPackage.contentTypeWorkbookDefault)
                {
                    Part.ContentType = ExcelPackage.contentTypeWorkbookDefault;
                }
            }
            else
            {
                if (Part.ContentType != ExcelPackage.contentTypeWorkbookMacroEnabled)
                {
                    Part.ContentType = ExcelPackage.contentTypeWorkbookMacroEnabled;
                }
            }
			
            UpdateDefinedNamesXml();

			// save the workbook
			if (_workbookXml != null)
			{
				_package.SavePart(WorkbookUri, _workbookXml);
			}

			// save the properties of the workbook
			if (_properties != null)
			{
				_properties.Save();
			}

            //Save the Theme
            ThemeManager.Save();

            // save the style sheet
            Styles.UpdateXml();
            _package.SavePart(StylesUri, this.StylesXml);
			
			SavePivotTableCaches();

			// save all the open worksheets
			var isProtected = Protection.LockWindows || Protection.LockStructure;
			foreach (var worksheet in Worksheets)
			{
				if (isProtected && Protection.LockWindows)
				{
					worksheet.View.WindowProtection = true;
				}
				worksheet.Save();
                worksheet.Part.SaveHandler = worksheet.SaveHandler;
			}

            // Issue 15252: save SharedStrings only once
            Packaging.ZipPackagePart part;
            if (_package.Package.PartExists(SharedStringsUri))
            {
                part = _package.Package.GetPart(SharedStringsUri);
            }
            else
            {
                part = _package.Package.CreatePart(SharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", _package.Compression);
                Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
            }

            part.SaveHandler = SaveSharedStringHandler;
            //UpdateSharedStringsXml();
			
			// Data validation
			ValidateDataValidations();

            //VBA
            if (_vba!=null)
            {
                VbaProject.Save();
            }

		}

        private void SavePivotTableCaches()
        {
			foreach (var info in _pivotTableCaches.Values)
			{
				foreach (var cache in info.PivotCaches)
				{
					if (cache._pivotTables.Count == 0)
					{
						cache.Delete();
					}
					//Rewrite the pivottable address again if any rows or columns have been inserted or deleted
					var r = cache.SourceRange;
					if (r != null)  //Source does not exist
					{
						ExcelTable t = null;
						if (r.IsName)
						{
							//Named range, set name
							cache.SetSourceName(((ExcelNamedRange)cache.SourceRange).Name);
						}
						else
						{
							var ws = Worksheets[cache.SourceRange.WorkSheetName];
							t = ws.Tables.GetFromRange(cache.SourceRange);
							if (t == null)
							{
								//Address
								cache.SetSourceAddress(cache.SourceRange.Address);
							}
							else
							{
								//Table, set name
								cache.SetSourceName(t.Name);
							}
						}

						var fields =
							cache.CacheDefinitionXml.SelectNodes(
								"d:pivotCacheDefinition/d:cacheFields/d:cacheField", NameSpaceManager);
						if (fields != null)
						{
							FixFieldNames(cache, t, fields);
							//cache.GenerateFieldSharedItems();
						}
						cache.CacheDefinitionXml.Save(cache.Part.GetStream(FileMode.Create));
					}
				}
			}
		}
		private void FixFieldNames(PivotTableCacheInternal cache, ExcelTable t, XmlNodeList fields)
		{
			int ix = 0;
			var flds = new HashSet<string>();
			cache.RefreshFields();
			foreach (XmlElement node in fields)
			{
				if (ix >= cache.SourceRange.Columns) break;
				var fldName = node.GetAttribute("name");                        //Fixes issue 15295 dup name error
				if (string.IsNullOrEmpty(fldName))
				{
					fldName = (t == null
						? cache.SourceRange.Offset(0, ix, 1, 1).Value.ToString()
						: t.Columns[ix].Name);
				}
				if (flds.Contains(fldName))
				{
					fldName = GetNewName(flds, fldName);
				}
				flds.Add(fldName);
				node.SetAttribute("name", fldName);
				cache.Fields[ix].WriteSharedItems(node, NameSpaceManager);
				ix++;
			}

		}
		private string GetNewName(HashSet<string> flds, string fldName)
		{
			int ix = 2;
			while (flds.Contains(fldName + ix.ToString(CultureInfo.InvariantCulture)))
			{
				ix++;
			}
			return fldName + ix.ToString(CultureInfo.InvariantCulture);
		}

		private void DeleteCalcChain()
		{
			//Remove the calc chain if it exists.
			Uri uriCalcChain = new Uri("/xl/calcChain.xml", UriKind.Relative);
			if (_package.Package.PartExists(uriCalcChain))
			{
				Uri calcChain = new Uri("calcChain.xml", UriKind.Relative);
				foreach (var relationship in _package.Workbook.Part.GetRelationships())
				{
					if (relationship.TargetUri == calcChain)
					{
						_package.Workbook.Part.DeleteRelationship(relationship.Id);
						break;
					}
				}
				// delete the calcChain part
				_package.Package.DeletePart(uriCalcChain);
			}
		}

		private void ValidateDataValidations()
		{
			foreach (var sheet in _package.Workbook.Worksheets)
			{
                if (!(sheet is ExcelChartsheet))
                {
                    sheet.DataValidations.ValidateAll();
                }
			}
		}

        private void SaveSharedStringHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
		{
            //Init Zip
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            stream.PutNextEntry(fileName);

            var cache = new StringBuilder();            
            var sw = new StreamWriter(stream);
            cache.AppendFormat("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count);
			foreach (string t in _sharedStrings.Keys)
			{

				SharedStringItem ssi = _sharedStrings[t];
				if (ssi.isRichText)
				{
                    cache.Append("<si>");
                    ConvertUtil.ExcelEncodeString(cache, t);
                    cache.Append("</si>");
				}
				else
				{
                    if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\r") || t.Contains("\t") || t.Contains("\n")))   //Fixes issue 14849
					{
                        cache.Append("<si><t xml:space=\"preserve\">");
					}
					else
					{
                        cache.Append("<si><t>");
					}
                    ConvertUtil.ExcelEncodeString(cache, ConvertUtil.ExcelEscapeString(t));
                    cache.Append("</t></si>");
				}
                if (cache.Length > 0x600000)
                {
                    sw.Write(cache.ToString());
                    cache = new StringBuilder();            
                }
			}
            cache.Append("</sst>");
            sw.Write(cache.ToString());
            sw.Flush();
            // Issue 15252: Save SharedStrings only once
            //Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
		}
		private void UpdateDefinedNamesXml()
		{
			try
			{
                XmlNode top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
				if (!ExistsNames())
				{
					if (top != null) TopNode.RemoveChild(top);
					return;
				}
				else
				{
					if (top == null)
					{
						CreateNode("d:definedNames");
						top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
					}
					else
					{
						top.RemoveAll();
					}
					foreach (ExcelNamedRange name in _names)
					{

						XmlElement elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
						top.AppendChild(elem);
						elem.SetAttribute("name", name.Name);
						if (name.IsNameHidden) elem.SetAttribute("hidden", "1");
						if (!string.IsNullOrEmpty(name.NameComment)) elem.SetAttribute("comment", name.NameComment);
						SetNameElement(name, elem);
					}
				}
				foreach (ExcelWorksheet ws in _worksheets)
				{
                    if (!(ws is ExcelChartsheet))
                    {
                        foreach (ExcelNamedRange name in ws.Names)
                        {
                            XmlElement elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
                            top.AppendChild(elem);
                            elem.SetAttribute("name", name.Name);
                            elem.SetAttribute("localSheetId", name.LocalSheetId.ToString());
                            if (name.IsNameHidden) elem.SetAttribute("hidden", "1");
                            if (!string.IsNullOrEmpty(name.NameComment)) elem.SetAttribute("comment", name.NameComment);
                            SetNameElement(name, elem);
                        }
                    }
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Internal error updating named ranges ",ex);
			}
		}

		private void SetNameElement(ExcelNamedRange name, XmlElement elem)
		{
			if (name.IsName)
			{
				if (string.IsNullOrEmpty(name.NameFormula))
				{
					if ((TypeCompat.IsPrimitive(name.NameValue) || name.NameValue is double || name.NameValue is decimal))
					{
						elem.InnerText = Convert.ToDouble(name.NameValue, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture); 
					}
					else if (name.NameValue is DateTime)
					{
						elem.InnerText = ((DateTime)name.NameValue).ToOADate().ToString(CultureInfo.InvariantCulture);
					}
					else
					{
						elem.InnerText = "\"" + name.NameValue.ToString() + "\"";
					}                                
				}
				else
				{
					elem.InnerText = name.NameFormula;
				}
			}
			else
			{
                elem.InnerText = name.FullAddressAbsolute;
			}
		}
		/// <summary>
		/// Is their any names in the workbook or in the sheets.
		/// </summary>
		/// <returns>?</returns>
		private bool ExistsNames()
		{
			if (_names.Count == 0)
			{
				foreach (ExcelWorksheet ws in Worksheets)
				{
                    if (ws is ExcelChartsheet) continue;
                    if(ws.Names.Count>0)
					{
						return true;
					}
				}
			}
			else
			{
				return true;
			}
			return false;
		}
		#endregion

		#endregion

		/// <summary>
		/// Removes all formulas within the entire workbook, but keeps the calculated values.
		/// </summary>
		public void ClearFormulas()
		{
			if (Worksheets == null || Worksheets.Count == 0) return;
			foreach(var worksheet in this.Worksheets)
			{
				worksheet.ClearFormulas();
			}
		}

		/// <summary>
		/// Removes all values of cells with formulas in the entire workbook, but keeps the formulas.
		/// </summary>
		public void ClearFormulaValues()
		{
			if (Worksheets == null || Worksheets.Count == 0) return;
			foreach (var worksheet in this.Worksheets)
			{
				worksheet.ClearFormulaValues();
			}
		}
		internal bool ExistsTableName(string Name)
		{
			foreach (var ws in Worksheets)
			{
                if (ws is ExcelChartsheet) continue;
                if (ws.Tables._tableNames.ContainsKey(Name))
				{
					return true;
				}
			}
			return false;
		}
		internal bool ExistsPivotTableName(string Name)
		{
			foreach (var ws in Worksheets)
			{
                if (ws is ExcelChartsheet) continue;
                if (ws.PivotTables._pivotTableNames.ContainsKey(Name))
				{
					return true;
				}
			}
			return false;
		}
		internal void AddPivotTableCache(PivotTableCacheInternal cacheReference)
		{			 
			CreateNode("d:pivotCaches");

			XmlElement item = WorkbookXml.CreateElement("pivotCache", ExcelPackage.schemaMain);
			item.SetAttribute("cacheId", cacheReference.CacheId.ToString());
			var rel = Part.CreateRelationship(UriHelper.ResolvePartUri(WorkbookUri, cacheReference.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
			item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

			var pivotCaches = WorkbookXml.SelectSingleNode("//d:pivotCaches", NameSpaceManager);
			pivotCaches.AppendChild(item);

			if(_pivotTableCaches.TryGetValue(cacheReference.SourceRange.FullAddress, out PivotTableCacheRangeInfo cacheInfo))
			{
				cacheInfo.PivotCaches.Add(cacheReference);
			}
			else
            {
				_pivotTableCaches.Add(cacheReference.SourceRange.FullAddress, new PivotTableCacheRangeInfo()
				{
					Address = cacheReference.SourceRange.FullAddress,
					PivotCaches=new List<PivotTableCacheInternal>() { cacheReference }
				});
			}
		}
		internal void RemovePivotTableCache(int cacheId)
		{
			string path = $"d:pivotCaches/d:pivotCache[@cacheId={cacheId}]";
			var relId = GetXmlNodeString(path + "/@r:id");
			DeleteNode(path,true);
			Part.DeleteRelationship(relId);
		}
		internal List<string> _externalReferences = new List<string>();
        //internal bool _isCalculated=false;
		internal void GetExternalReferences()
		{
			XmlNodeList nl = WorkbookXml.SelectNodes("//d:externalReferences/d:externalReference", NameSpaceManager);
			if (nl != null)
			{
				foreach (XmlElement elem in nl)
				{
					string rID = elem.GetAttribute("r:id");
					var rel = Part.GetRelationship(rID);
					var part = _package.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
					XmlDocument xmlExtRef = new XmlDocument();
                    LoadXmlSafe(xmlExtRef, part.GetStream()); 

					XmlElement book=xmlExtRef.SelectSingleNode("//d:externalBook", NameSpaceManager) as XmlElement;
					if(book!=null)
					{
						string rId_ExtRef = book.GetAttribute("r:id");
						var rel_extRef = part.GetRelationship(rId_ExtRef);
						if (rel_extRef != null)
						{
							_externalReferences.Add(rel_extRef.TargetUri.OriginalString);
						}

					}
				}
			}
		}
        /// <summary>
        /// Disposes the workbooks
        /// </summary>
        public void Dispose()
        {
            if (_sharedStrings != null)
            {
                _sharedStrings.Clear();
                _sharedStrings = null;
            }
            if (_sharedStringsList != null)
            {
                _sharedStringsList.Clear();
                _sharedStringsList = null;
            }
            _vba = null;
            if (_worksheets != null)
            {
                _worksheets.Dispose();
                _worksheets = null;
            }
            _package = null;
            _properties = null;
            if (_formulaParser != null)
            {
                _formulaParser.Dispose();
                _formulaParser = null;
            }   
        }

        internal void ReadAllTables()
        {
            if (_nextTableID > 0) return;
            _nextTableID = 1;
            _nextPivotTableID = 1;
            foreach (var ws in Worksheets)
            {
                if (!(ws is ExcelChartsheet)) //Fixes 15273. Chartsheets should be ignored.
                {
                    foreach (var tbl in ws.Tables)
                    {
                        if (tbl.Id >= _nextTableID)
                        {
                            _nextTableID = tbl.Id + 1;
                        }
                    }
                    foreach (var pt in ws.PivotTables)
                    {
                        if (pt.CacheId >= _nextPivotTableID)
                        {
                            _nextPivotTableID = pt.CacheId + 1;
                        }
                    }
                }
            }
        }

		internal Dictionary<string, ExcelSlicerCache> _slicerCaches = null;
		internal ExcelSlicerCache GetSlicerCaches(string key)
        {
			if(_slicerCaches==null)
            {
				_slicerCaches = new Dictionary<string, ExcelSlicerCache>();
				foreach (var r in Part.GetRelationshipsByType(ExcelPackage.schemaRelationshipsSlicerCache))
				{
					var p = Part.Package.GetPart(UriHelper.ResolvePartUri(WorkbookUri, r.TargetUri));
					var xml = new XmlDocument();
					LoadXmlSafe(xml, p.GetStream());

					ExcelSlicerCache cache;
					if(xml.DocumentElement.FirstChild.LocalName=="pivotTables")
					{
						cache = new ExcelPivotTableSlicerCache(NameSpaceManager);
					}
					else
                    {
						cache = new ExcelTableSlicerCache(NameSpaceManager);
					}
					cache.CacheRel = r;
					cache.Part = p;
					cache.TopNode = xml.DocumentElement;
					cache.Init(this);

					_slicerCaches.Add(cache.Name, cache);
				}
			}
			if (_slicerCaches!=null && _slicerCaches.TryGetValue(key, out ExcelSlicerCache c))
			{
				return c;
			}
			else
            {
				return null;
            }
        }

        internal ExcelTable GetTable(int tableId)
        {
            foreach(var ws in Worksheets)
            {
				var t=ws.Tables.FirstOrDefault(x => x.Id == tableId);
				if(t!=null)
                {
					return t;
                }
            }
			return null;
        }
    } // end Workbook
}
