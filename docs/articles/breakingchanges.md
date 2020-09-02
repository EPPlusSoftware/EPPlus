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
