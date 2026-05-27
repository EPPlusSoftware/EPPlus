## Worksheets collection behavior in .NET Framework
The default behavior for the Worksheet collection base in .NET Framework has changed from 1 to 0. 
This is the same behavior as in .NET Core.
For backward compatibility the property<code>IsWorksheets1Based</code> can be set to true, to have the same behavior as in previous version of EPPlus.
This property can also be set in the app.config:

```xml
  <appSettings>
   <!-- Set worksheets collection to start from one.Default is 0,  
        for backward compatibility reasons this can be altered 1. -->  
   <add key = "EPPlus:ExcelPackage.Compatibility.IsWorksheets1Based" value="true" />
   </appSettings>
```

### Moved from Global Namespace
* eShapeStyle --> OfficeOpenXml.Drawing
* eTextAlignment --> OfficeOpenXml.Drawing
* eFillStyle --> OfficeOpenXml.Drawing
* eEndStyle --> OfficeOpenXml.Drawing
* eEndSize --> OfficeOpenXml.Drawing

### Package
Package property of ExcelPackage has been removed.

### Styles
Unused ExcelXsf.Styles has been removed 

### Named ranges
Misspelled method `AddFormla` has been removed. Use `AddFormula`

### Theme
The `Theme` property on `ExcelColor` (used on cell styles) has changed datatype from string to the enum `eThemeSchemeColor`

### Drawing
Drawings will now move and size when inserting/deleting rows/columns depending on the `ExcelDrawing.EditAs` property
ExcelDrawings UriDrawing from public to internal

### Picture
Pictures have changed the behavior as the oneCellAnchor tag is used instead of the twoCellAnchor tag with the editAs="oneCell".
This will generate an Exception if you try to set the editAs property to anything but "oneCell".
The old behavior can be set via the Picture.ChangeCellAnchor method.

### Chart
`ExcelScatterChartSerie`, `ExcelLineChartSerie` och `ExcelRadarChartSerie` has changed the datatype of the `Marker` Property from the enum `eMarkerStyle` to a new `Marker` class.
The old `Marker` property can be set via the new class `Marker.Style`
ExcelChart has been changed to an abstract class from Version 5.2. Standard charts has a new implementation class called ExcelChartStandard. The new Extended charts uses the ExcelChartEx class.
A chart that reference within its own worksheet will now change the worksheet in the series addresses for any copy made with the Worksheets.Add method

### Formula parser
Handling of circular references has been redesigned to better reflect Excel. 
Changed misspelled property name
Misspelled property `ExcelCalculationOption.AllowCirculareReferences` has been removed. Please use `ExcelCalculationOption.AllowCircularReferences`

### Pivot tables
Misspelled property `ColumGrandTotals` has been removed. Please use `ColumnGrandTotals`
Pivot tables will always have the flag to be refreshed on load set.
Pivot table filter classes moved to correct namespace --> OfficeOpenXml.Table.PivotTable

### Breaking changes from EPPlus 5.6
Inserting rows in tables will by default copy any style from the row before. 
The ExcelRange.Cells indexer will not permit accessing other worksheets using the string address overload (for example sheet1.Cells["sheet2.A1"]).

### Breaking changes from EPPlus 5.8
LoadFromCollection changes the data type of parameter 'TableStyle' from TableStyles to TableStyles?. 
The default value, if omitted, changes from TableStyles.None to null. TableStyles.None, if supplied will create a table with style None.

### Breaking changes from EPPlus 6.0
Targeting framework for .NET4.0 has been removed. 
Targeting framework for .NET 4.5 has been upgraded to .NET 4.52.
All references to System.Drawing.Common has been removed. See [Breaking Changes in EPPlus 6](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-6) for more information.
Static class 'FontSize' has splitted width and heights into two dictionaries. FontSizes are lazy-loaded when needed. 

### Breaking changes from EPPlus 6.2
Updating data validations via the XML DOM will not work as read and write is performed on load/save. ExcelDataValidation.IsStale is deprecated and will always return false.

### Breaking changes from EPPlus 7.0
The formula parser has changed significantly in EPPlus 7, requiring all custom functions that are inherited from the `ExcelFunction` class to be reviewed. 
The `ExcelFunction` class has changed, now exposes new properties used to handle array results and condition behaviour. 
* The `Execute` method has changed signature changing the `IEnumarable` in the first parameter to `IList`. New signature is: Execute(IList, ParsingContext).
* `ArgumentMinLength` - Required. Minimum number of parameters supplied to the function. Suppling less parameters to the function will result in a #VALUE! error.
* `NamespacePrefix` - If the function requires a prefix when saved, for example "_xlfn." or "_xlfn._xlws."
* `HasNormalArguments` A Boolean indicating if the formula only has normal arguments. If false, the `GetParameterInfo` method must be implemented. The default is true.
* `ReturnsReference` - If true the function can return a reference to a range. Use the `CreateAddressResult` to return the result with a reference. Returning a reference will cause the dependency chain to check the address and will allow the colon operator to be used with the function.
* `IsVolatile` - If the function returns a different result when called with the same parameters. The default is false.
* `ArrayBehaviour` - If the function allows arrays as input in a parameter, resulting in an array output. Also see the `GetArrayBehaviourConfig` method.
* IFunctionModules.CustomCompilers has been removed and compilers can no longer be added. This has been replaced by ExcelFunction.ParameterInfo and ExcelFunction.ArrayBehaviour which configures the new behaviour of the formula calculation engine.
* `CalculateCollection` - has been removed. EPPlus no long uses collections of FunctionArgument in this way. Use the `InMemoryRange` class instead.
* Converting double's from strings in the formula parser will now use try to parse the string using the the CurrentCulture before trying the InvariantCulture.
* The default value of `ExcelCalculationOption.PrecisionAndRoundingStrategy` in the formula calculation has been changed from `DotNet` to `Excel`
* The `ErrorHandlingFunction` class has been removed. Use the ParametersInfo property with  FunctionParameterInformation.IgnoreErrorInPreExecute instead.
#### Methods
* `CreateAddressResult`  - Returns the result with a reference to a range.
* `CreateDynamicArrayResult` - The result should be treated as a dynamic array.
* `GetArrayBehaviourConfig` - Sets the index if the parameters that can be arrays. Also see the `ArrayBehaviour` property.
The ExcelFunctionArgument class
 * The `GetAsRangeInfo(ParsingContext)` has been removed. Use the `ValueAsRangeInfo` property instead.
 * The `IsEnumerableOfFuncArgs` and `ValueAsEnumerableOfFuncArgs` properties and  has been removed.
 * The `SetExcelStateFlag` and `ExcelStateFlagIsSet` methods has been removed.
Misspelled property `ExcelIgnoreError.CalculatedColumm` has been renamed `CalculatedColumn`
#### Tokenizer, Expressions and Compile result
* The source code tokenizer now tokenizes in more detail, tokenizing addresses. 
* The expression handling is totally rewritten and now uses reversed polish notation instead of an expression tree. This change affects internal classes only.
* The `CompileResult` class has moved to a new namespace: OfficeOpenXml.FormulaParsing.FormulaExpressions
* Adding defined names referencing addresses will now be added as fixed addresses (i.e $A$1), unless the `allowRelativeAddress` parameter of the `ExcelNamedRangeCollection.Add` method is set to true.
* `ParsingConfiguration.Lexer` and `ParsingConfiguration.SetLexer(ILexer lexer)` has been removed.
* `ParsingConfiguration.SetExpresionCompiler` has been removed. 
#### Autofilter & Table filter
*  Added `ExcelPackageSettings.ApplyFiltersOnSave` to decide if Filters will be applied on Saving the workbook. Default is True. If set to false, you will call the `ApplyFilter` method manually to show/hide rows the matches the filters criterias.
*  `ExcelWorksheet.AutofilterAddress` is now obsolete. Use `ExcelWorksheet.Autofilter.Address` instead. `ExcelWorksheet.Autofilter` will now always be set instead of being null if no autofilter was present.
#### ConditionalFormatting
* Updating ConditionalFormatting via the XML DOM will not work as read and write is performed on load/save.
* The base class `ConditionalFormattingRule` and all derived classes no longer contain the Node property.
* Misspelled enum member `eTrendLine.MovingAvgerage` has been removed and replaced with `eTrendLine.MovingAverage`
* ConditionalFormatting classes are now Internal. Interfaces for each class exist and have all relevant properties instead.
#### ExcelHyperlink
* Renamed misspelled properties `ColSpann` and `RowSpann` to `ColSpan` and `RowSpan` on the `ExcelHyperLink` class.

### Breaking changes from EPPlus 7.1
*Misspelled property `MemorySettings.MemoryManger` was renamed `MemoryManager`
#### Defined Names
* EPPlus will now encode string values and in defined name .
#### Data Validation
* Removed DataValidationStaleException as DataValidations cannot be stale since Epplus 7.
#### Conditional Formatting
* When reading conditional formatting from file Style.Fill.PatternType is now always null if the patternType attribute in the xml has not been set.
#### Rich Text
* The `ExcelRichText._collection` has been set to internal.
* Public static methods in the class `XmlHelper` used to getting richtext properties has been changed to internal.
* Table Column Names
	* ShowHeaders = True property on tables no longer causes crash in rare cases. It also no longer updates column names.
	* Table.SyncColumnNames method added to ensures column names and cell-values in header are equal. Applying this method should cover any potential issues caused by above fix not updating column names.
	* Adding a table column to a table no longer creates a column name that can conflict with existing names.

### Breaking changes from EPPlus 7.2
* Changed class ExcelTextFormatBase to abstract
* OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelReferenceType RelativeRowAbsolutColumn corrected to RelativeRowAbsoluteColumn

### Breaking changes from EPPlus 7.4
* Removed unused class ParsingScope and ParsingScopes.
* Removed unused interface IParsingLifetimeEventHandler and implemetation.
* Removed implementation of IParsingLifetimeEventHandler.ParsingCompleted in the ParsingContext class.

### Breaking change from EPPlus 7.5.2
Renaming worksheet's will now change the formula correctly to include single quotes for the worksheet name if necessary.

### Breaking changes from EPPlus 7.6.0
* Altering the worksheets collection when iterating it using IEnumerable, now throws an InvalidOperationException.

### Breaking changes from EPPlus 8.0
* Set ExcelPackageSettings.ApplyFiltersOnSave default value to false.
* RichText now returns font name, size and font family from cell style if not set.
* Fixed spelling error in ExcelDrawingGradientFillLinearSettings. `public double Angel` is now `public double Angle`.
* Removed reference to EPPlus.System.Drawing for primary image and text handlers. The generic handler is now used for all target frameworks. 
  The 'SystemDrawingTextMeasurer. can still be used referencing the EPPlus.System.Drawing nuget package and set the 'PrimaryTextMeasurer' or the 'PrimaryImageHandler'.
* Switched default theme to the newest Excel theme (202300)
  You can still use these handlers by referencing the EPPlus.System.Drawing handlers nuget package and use the 'SystemDrawingTextMeasurer' or 'SystemDrawingImageHandler' classes as primary handler.
  Also see https://github.com/EPPlusSoftware/EPPlus/wiki/Autofit-columns.
* Adding Threaded Comments now throws an exception if the 'personId' does not exist in the ThreadedCommentPersons Collection.
* Worksheet.Dimension now returns the address including rows and columns data.
* Removed enum OfficeOpenXml.FormulaParsing.Excel.Operators.LimitedOperators as it was unused.

#### Removed Methods & Properties
* Obsolete property `ExcelVbaReferenceControl.LibIdExternal`, please use `LibIdExtended` instead.
* Obsolete property `ExcelDataValidation.IsStale` has been removed.
* Obsolete property `ExcelLineChartSerie.LineColor` has been removed. Use `Border.Fill.Color` instead.
* Obsolete property `ExcelLineChartSerie.MarkerSize` has been removed. Use `Marker.Size` instead.
* Obsolete property `ExcelLineChartSerie.LineWidth` has been removed. Use `Border.Width` instead.
* Obsolete property `ExcelLineChartSerie.MarkerLineColor` has been removed. Use `Marker.Border.Fill.Color` instead.
* Obsolete property `ExcelRadarChartSerie.MarkerSize` has been removed. Use `Marker.Size` instead.
* Obsolete property `ExcelScatterChartSerie.LineColor` has been removed. Use `Border.Fill.Color` instead.
* Obsolete property `ExcelScatterChartSerie.MarkerSize` has been removed. Use `Marker.Size` instead.
* Obsolete property `ExcelScatterChartSerie.MarkerColor` has been removed. Use `Marker.Fill` instead.
* Obsolete property `ExcelScatterChartSerie.LineWidth` has been removed. Use `Border.Width` instead.
* Obsolete property `ExcelScatterChartSerie.MarkerLineColor` has been removed. Use `Marker.Border.Fill.Color` instead.
* Obsolete property `ExcelStockChartSerie.LineColor` has been removed. Use `Border.Fill.Color` instead.
* Obsolete property `ExcelStockChartSerie.MarkerSize` has been removed. Use `Marker.Size` instead.
* Obsolete property `ExcelStockChartSerie.MarkerColor` has been removed. Use `Marker.Fill` instead.
* Obsolete property `ExcelStockChartSerie.LineWidth` has been removed. Use `Border.Width` instead.
* Obsolete property `ExcelStockChartSerie.MarkerLineColor` has been removed. Use `Marker.Border.Fill.Color` instead.
* Obsolete Method `ExcelDrawings.AddPicture(string, Stream, ePictureType?)` has been removed. Use `ExcelDrawings.AddPicture(string, Stream)` instead.
* Obsolete Method `ExcelDrawings.AddPicture(string, Stream, ePictureType?, Uri)` has been removed. Use `ExcelDrawings.AddPicture(string, Stream, Uri)` instead.
* Obsolete property `ExcelHeaderFooter.InsertPicture` has been removed.
* Obsolete Constructor `ExcelHyperLink(string, bool)` has been removed. Depricated constructor, use `ExcelHyperLink(string)` instead
* Obsolete property `ExcelStockChartSerie.MarkerColor` has been removed. Use `Marker.Fill` instead.
* Obsolete property `ExcelRow.RowID` has been removed.
* Obsolete interface `IRangeID` has been removed.
* Obsolete property `ExcelWorksheet.AutoFilterAddress` has been removed. Use `AutoFilter.Address` instead.
* Obsolete Method `ExcelWorksheet.DeleteRow(int, int, bool)` has been removed. Use `ExcelWorksheet.DeleteRow(int, int)` instead.
* Obsolete Method `ExcelFunction.ValidateArguments(IEnumerable<FunctionArgument>, int, eErrorType)` has been removed. Use property `ArgumentMinLength` instead.
* Obsolete Method `ExcelFunction.ValidateArguments(IEnumerable<FunctionArgument>, int)` has been removed. Use property `ArgumentMinLength` instead.
* Obsolete property `ExcelPivotTable.TableStyle` has been removed. Use property `PivotTableStyle ` instead.

#### Changed data type
`ExcelChartTrendline.Order` from `decimal` to `double`.
`ExcelChartTrendline.Period` from `decimal` to `double`.
`ExcelChartTrendline.Forward` from `decimal` to `double`.
`ExcelChartTrendline.Backward` from `decimal` to `double`.
`ExcelChartTrendline.Intercept` from `decimal` to `double`.
`ExcelDoughnutChart.FirstSliceAngle` from `decimal` to `double`.
`ExcelDoughnutChart.HoleSize from` `decimal` to `double`.
`ExcelTime` constructor now takes `double` as parameter instead of `decimal`.
`ExcelTime.ToExcelTime()` returns a `double` instead of `decimal`.
`ExcelWorkbook.MaxFontWidth` from `decimal` to `double`.
`ExcelView3D.Perspective` from `decimal` to `double`.
`ExcelView3D.RotX` from `decimal` to `double`.
`ExcelView3D.RotY` from `decimal` to `double`.
`ExcelPrinterSettings.LeftMargin` from `decimal` to `double`.
`ExcelPrinterSettings.RightMargin` from `decimal` to `double`.
`ExcelPrinterSettings.TopMargin` from `decimal` to `double`.
`ExcelPrinterSettings.BottomMargin` from `decimal` to `double`.
`ExcelPrinterSettings.HeaderMargin` from `decimal` to `double`.
`ExcelPrinterSettings.FooterMargin` from `decimal` to `double`.
`ExcelSparklineColor.Tint` from `decimal` to `double`.
`ExcelColorXml.Tint` from `decimal` to `double`.
`ExcelColor.Tint` from `decimal` to `double`.

### 8.0.2
* Removed base class from ExcelVmlDrawingPosition and with that the Load, UpdateXml methods and the RowOff and ColOff properties as they were duplicates.
### 8.5.0
* Setting dataLabelPosition.Top on BarCharts corrupted the excel file when saved. Trying to set this now throws an error instead. 
* NumberFormatToTextArgs.NumberFormat now returns the interface IExcelNumberFormat rather than the ExcelNumberFormatXml class. All public variables remain the same and it can be safely cast to ´ExcelNumberFormatXml´ as long as ´Package.Workbook.NumberFormatToTextHandler´ is null.


### 8.5.5
* `ws.Cells["A1"].RichText` no longer sets cells with `null` to `string.empty`
* .RichText no longer sets the cell or contents to be RichText automatically.
 This is instead done when properties such as; `.Text`, `.Add` or `.Insert` are set on the .RichText property.
The misspelled enum eCompundLineStyle has been renamed eCompoundLineStyle.


### 9.0.0
The misspelled enum eCompundLineStyle has been renamed eCompoundLineStyle.