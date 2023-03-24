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
