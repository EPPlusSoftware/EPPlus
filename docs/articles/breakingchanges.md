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

### Breaking Change From EPPlus 5.6
Inserting rows in tables will by default copy any style from the row before. 
The ExcelRange.Cells indexer will not permit accessing other worksheets using the string address overload (for example sheet1.Cells["sheet2.A1"]).

### Breaking Change From EPPlus 5.8
LoadFromCollection changes the data type of parameter 'TableStyle' from TableStyles to TableStyles?. 
The default value, if omitted, changes from TableStyles.None to null. TableStyles.None, if supplied will create a table with style None.

### Breaking Change From EPPlus 6.0
Targeting framework for .NET4.0 has been removed. 
Targeting framework for .NET 4.5 has been upgraded to .NET 4.52.
All references to System.Drawing.Common has been removed. See [Breaking Changes in EPPlus 6](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-6) for more information.
Static class 'FontSize' has splitted width and heights into two dictionaries. FontSizes are lazy-loaded when needed. 

### Breaking Change From EPPlus 6.2
Updating data validations via the XML DOM will not work as read and write is performed on load/save. ExcelDataValidation.IsStale is deprecated and will always return false.

### Breaking Change From EPPlus 7.0
The formula parser has changed significantly in EPPlus 7, requiring all custom functions that are inherited from the `ExcelFunction` class to be reviewed. 
The `ExcelFunction` class has changed, now exposes new properties used to handle array results and condition behaviour. 
* The `Execute` method has changed signature changing the `IEnumarable` in the first parameter to `IList`. New signature is: Execute(IList, ParsingContext).
* `ArgumentMinLength` - Required. Minimum number of parameters supplied to the function. Suppling less parameters to the function will result in a #VALUE! error.
* `NamespacePrefix` - If the function requires a prefix when saved, for example "_xlfn." or "_xlfn._xlws."
* `HasNormalArguments` A Boolean indicating if the formula only has normal arguments. If false, the `GetParameterInfo` method must be implemented. The default is true.
* `ReturnsReference` - If true the function can return a reference to a range. Use the `CreateAddressResult` to return the result with a reference. Returning a reference will cause the dependency chain to check the address and will allow the colon operator to be used with the function.
* `IsVolatile` - If the function returns a different result when called with the same parameters. The default is false.
* `ArrayBehaviour` - If the function allows arrays as input in a parameter, resulting in an array output. Also see the `GetArrayBehaviourConfig` method.
#### Methods
* `CreateAddressResult`  - Returns the result with a reference to a range.
* `CreateDynamicArrayResult` - The result should be treated as a dynamic array.
* `GetArrayBehaviourConfig` - Sets the index if the parameters that can be arrays. Also see the `ArrayBehaviour` property.
Misspelled property `ExcelIgnoreError.CalculatedColumm` has been renamed `CalculatedColumn`
#### Tokenizer, Expressions and Compile result
* The source code tokenizer now tokenizes in more detail, tokenizing addresses. 
* The expression handling is totally rewritten and now uses reversed polish notation instead of an expression tree. This change affects internal classes only.
* The `CompileResult` class has moved to a new namespace: OfficeOpenXml.FormulaParsing.FormulaExpressions
* Adding defined names referencing addresses will now be added as fixed addresses (i.e $A$1), unless the `allowRelativeAddress` parameter of the `ExcelNamedRangeCollection.Add` method is set to true.
#### ConditionalFormatting
* Updating ConditionalFormatting via the XML DOM will not work as read and write is performed on load/save.
* The base class `ConditionalFormattingRule` and all derived classes no longer contain the Node property.
