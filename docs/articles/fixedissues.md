﻿# Features / Fixed issues - EPPlus 6

## Version 6.1.3
### Fixed issues
* When clearing a formula and then insert a row into the worksheet an exception was thrown.
* Having a pivot field with string grouping caused an Exception on loading a workbook.
* ExcelExternalWorkbook.UpdateCache() throw an NullReferenceException if a worksheet name did not exist.
* Applying a style for a worksheet that has a style set on the column level, did not retain the style for the last columns in some cases.
* Could not delete the last row (1048576) in a worksheet.
* Group drawings did not update children when rows were resized.Thanks to gnoul-mah for pr.

## Version 6.1.2
### Fixed issues
* Fixed an issue with the Roman function. Thanks to ihorbach.
* Fixed a performance issue with calculated table columns.
* Having hyperlinks longer than 2079 characters resulted in a corrupt workbook. EPPlus will now crop hyperlinks over 2079 characters
* Date functions Month, Day, Hour,Second and DateDiff used InvariantCulture instead of CurrentCulture.
* Checkboxes in the style dialog for named styles was not retained after a workbook was saved with EPPlus.Thanks to ihorbach.
* Iterating over a range with multiple comma-separated ranges iterated the first range twice.
* Creating Array Formulas via the ExcelRange.CreateArrayFormula did not create the formula correctly.
* Custom row heights was not copied correctly.

## Version 6.1.1
### Minor features
* Add support for linking a cell to a chart title text
### Fixed issues
* EPPlus did not preserve "Host Extender Info" in the vba project.
* VBA signing failed on new workbooks.
* ExcelTable sorting corrupted relative formulas.
* Images in a group did not change position when row height changed.
* RichText.Color and ExcelColor.LookupColor did not return the correct color if the color was Auto or Theme and did not always adjust for Tint.
* Workbooks with shared strings without a reference and a blank v-element failed to load.
### Other changes
* Added Target framework .NET 7.

## Version 6.1.0
### Features
* Support for new types of VBA signing. See [This link](https://github.com/EPPlusSoftware/EPPlus/wiki/VBA)
	* Agile VBA Signing
	* V3 VBA Signing
* Change hash algorithm on a VBA signature. Supports MD5, SHA1, SHA256, SHA384 and SHA512.
* ExcelRange.ToCollection method - to map ranges and tables to an collection of T. See [This link](https://github.com/EPPlusSoftware/EPPlus/wiki/ToCollection)
* New methods to group, ungroup, collapse, and expand rows and columns - See [This link](https://github.com/EPPlusSoftware/EPPlus/wiki/Grouping-and-Ungrouping-Rows-and-Columns)
	* Group method method
	* Ungroup method method
	* CollapseChildren method
	* ExpandChildren method
	* SetVisibleOutlineLevel method
* New overloads of Drawings.AddPicture that reads the signature of the image from stream to identify the type of image.
	* AddPicture(string, Stream)
	* AddPicture(string, Stream, Uri)
These overloads have been deprecated:
	* AddPicture(string, Stream, ePictureType)
	* AddPicture(string, Stream, ePictureType, Uri) 
* Preserve the 'Table' formula properties (Created via the What-If Analysis-Data Table).
### Fixed issues
* Invalid formatted hyperlinks were not loaded and saved correctly.
* Rows to repeat is not copied when adding a worksheet with another worksheet as template.
* Rows with the default height got an incorrect height when copied to a new worksheet if the Normal style had a font other than the default.
* The GenericImageHandler failed to load images in .NET in Unity as System.Drawing was not available.
* EpplusTableColumnAttribute NumberFormat is ignored when Header contains space.
* Workbooks with color styles not having any attributes, failed to load.
### Other changes
* Target framework .NET 5.0 has been removed as it is out of support by Microsoft.

## Version 6.0.8
### Fixed issues
* Fixed issue with the DataValidations.Clear() method, where DataValidations were only removed from the collection classes in EPPlus but not removed from the underlying xml.
* Fixed issue where inserting/deleting columns in the source of a pivot table sometimes caused the PivotTable to become corrupt.
* Fixed issue with Png files without the pHYs chunk that failed to Add to the Drawings collection.
* Adjusted the EPPlus source code to avoid validation errors from the PEVerify tool
* Added error handling to the initialization of RecyclableMemoryStreamManager which previously caused an uncontrolled Exception on the Unity platform
* When saving a workbook with external references to another workbook, EPPlus now updates the reference with a relative path instead of an absolute path.

## Version 6.0.7
### Minor features
* New static method Configure on ExcelPackage to configure location of config files and error handling. See  [This link...](https://github.com/EPPlusSoftware/EPPlus/wiki/Configuration)
### Fixed issues
* Copying comments sometimes did not change the name in the vml drawing causing OOXML validation to complain on duplicate drawings.
* The count attribute on 'xsf' node in Styles.xml was not correctly set.
* Setting drawing coordinates did not update the xml on save if ExcelPackage.DoAdjustDrawings was set to false.
* IF function now handle errors correctly.
* Added method ExcelRangeBase.SetErrorValue to set a cell to an error value. Added static methods ExcelErrorValue.Create and ExcelErrorValue.Parse.
* Referencing a worksheet to a cell address after a colon (for example 'sheet1'!a1:'sheet1'!A3) did not work correctly in the formula calculation.
* Added RichText and TextBody properties to ExcelControlWithText.
* Removing / Clearing or inserting into a table with a Calculated Column Formula sometimes caused a corrupt workbook or an Exception. 
* Removing VBA signtures did not remove the newer types of signatures, Agile and V3.
* ExcelRangeBase.ToDataTable could not export a range with no column names and a single row.

## Version 6.0.6
### Fixed issues
* Fixed an error in Positioning and sizing of form controls intruduced in 5.8.12.
* Pivot table styles in template workbooks sometimes corrupts the new workbook.
* Spaces were not preserved in rich text in drawing objects.
* Defined names referencing external reference sometimes loaded incorrectly.
* Drawing.ChangeCellAnchor causes a corrupt package in some cases

## Version 6.0.5
### Fixed issues
* VBA code modules with LF only as EOL, causes code module to load incorrectly.
* INDIRECT function did not always set the correct data type when returning a single cell.
* Clearing ranges with threaded comments caused an exception.
* Copying drawings with hyperlinks failed.
* Improves insert/delete performance when working with drawings.
* Insert row caused images to change size in some cases when having two anchored drawings.
* Improved handling of percentage values in strings in formula calculation.
* Custom index colors are now loaded from the styles.xml, thanks to Raboud.
* Fix to get the app0 header length correct for internal jpg reader for the internal image reader.
* Images in the header row was not correctly rendered in the HTML export for ranges.
* Upgraded .NET 4.5.2 to 4.6.2, as 4.5.2 has reached End of Support

## Version 6.0.4
### Fixed issues
* Improved handling of circular references for SUMIF and COUNTIF
* Memory optimization reading directly from the zip stream, when reading the worksheet xml, allowing unexteracted parts larger than 2GB.
* Hyperlinks in cells added with 'Display=null' will now use the formatted cell value as text for in workbook links.
* Remove invalid attribute TopLeftCell on the Selection element for splited/freezed worksheets.
* Fixed an issue in the unziping of packages using ZIP64 file headers and data descriptors
* Fix for removing rows from ExcelTable with options ShowHeader set to false.
* Improved handling of workbook- and worksheet-names when loaded from file with a relative address
* Box and Whisker chart series failed when copied to a new worksheet.

## Version 6.0.3
### Changes
### Features
* Html Export for tables and ranges, See [Html Export](https://github.com/EPPlusSoftware/EPPlus/wiki/HTML-Export)
* Json Export for tables and ranges, See [Json Export](https://github.com/EPPlusSoftware/EPPlus/wiki/JSON-Export)
* Breaking Change: Removed System.Drawing.Common from all public classes, See [Breaking Changes in EPPlus 6](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-6) for more information
* 32 new functions:
	* BETADIST	
	* BETA.DIST
	* BETAINV
	* BETA.INV
	* CHIDIST
	* CHISQ.DIST.RT
	* CHIINV
	* CHISQ.INV
	* CHISQ.INV.RT
	* CORREL
	* EXPONDIST
	* EXPON.DIST
	* FISHER
	* FISHERINV
	* FORECAST
	* FORECAST.LINEAR
	* GAUSS
	* GEOMEAN
	* HARMEAN
	* INTERCEPT
	* KURT
	* PEARSON
	* PHI
	* RSQ
	* SKEW
	* SKEW.P
	* STANDARDIZE
	* ACCRINT
	* ACCRINTM
	* TBILLEQ
	* TBILLPRICE
	* TBILLYIELD

### Minor Features
* Breaking Change: Static class 'FontSize' has splitted width and heights into two dictionaries. FontSizes are lazy-loaded when needed. 
* New ExcelRangeBase.GetCellValue<T> method
* New overload for ExcelRangeBase.LoadFromDictionaries method with IEnumerable<dynamic>.
* Added Datatypes and Culture to LoadFromDictionariesParams. This is the settings for the ExcelRangeBase.ExcelRangeBase.LoadFromDictionaries.
* Added ExcelRichTextCollection.HtmlText and ExcelRichText.HtmlText property.

## Version 5.8.9
### Fixed issues
* Fixed issue with start_num parameter for functions FIND and SEARCH
* Pivot table slicers in a template sometimes caused a corrupt workbook on save
* Pivot table fields that had subtotals and null values in shared cache items caused the package to fail on load in some cases.
* Having the value set to 0(zero) and the number format to date or time returned the format instead of the formatted value.
* DeleteColumn caused the worksheet to expands to the maximum column properties extended to the last column (XFD).
* The UPPER and LOWER functions did not handle empty cell values correctly.
* Fixed an issue in ExpressionEvaluator when evaluating empty string criterias.

## Version 5.8.10
### Fixed issues
* Hyperlinks in cells added with 'Display=null' will now use the formatted cell value as text for in workbook links.
* Remove invalid attribute TopLeftCell on the Selection element for splited/freezed worksheets.
* Fixed an issue in the unziping of packages using ZIP64 file headers and data descriptors
* Fix for removing rows from ExcelTable with options ShowHeader set to false.
* Improved handling of workbook- and worksheet-names when loaded from file with a relative address
* Box and Whisker chart series failed when copied to a new worksheet.

## Version 5.8.9
### Fixed issues
* Fixed issue with start_num parameter for functions FIND and SEARCH
* Pivot table slicers in a template sometimes caused a corrupt workbook on save
* Pivot table fields that had subtotals and null values in shared cache items caused the package to fail on load in some cases.
* Having the value set to 0(zero) and the number format to date or time returned the format instead of the formatted value.
* DeleteColumn caused the worksheet to expands to the maximum column properties extended to the last column (XFD).
* The UPPER and LOWER functions did not handle empty cell values correctly.
* Fixed an issue in ExpressionEvaluator when evaluating empty string criterias.


## Version 5.8.8
### Fixed issues
* Removed unnessesary Nuget references to packages already included in the targeting frameworks.
* Fixed calculation issue when ExcelPackage.Compatibility.IsWorksheets1Based is set to true
* Added new method GetTable() and the propert 'IsTable' to ExcelRangeBase to get the table object if the range corresponds with the tables Range.

## Version 5.8.7
### Fixed issues
* LoadAsync(FileInfo) and LoadAsync(string) do not close the file stream.
* Hyperlinks referencing the same relation fails to load the package.
* Monochromatic chart color schemes gave the wrong tint/shade for multiple series.
* Added missing property 'JustifyLastLine' to ExcelStyle and ExcelXfsXml. Collapsed do not set Hidden on columns when set to true.
* Improved handling of defined names in range addresses in formula calc, for example SUM("MyRange1:MyRange2") 
* Escaped double quotes did not work properly for shared formulas in formula calc.
* Table.AddRow did not adjust Conditional Formatting and Data Validation.
* Support for numeric criteria with operators in MINIFS/MAXIFS functions.
* ExcelWorksheet.ClearFormulas method throw a NullReferenceException if ExcelWorksheet.Dimension was null.

## Version 5.8.6
### Fixed issues
* Rounding functions now returns 0 when referencing empty cells
* Copying elements in vml files caused attributes to lose there name space and create a duplicate.
* Removed most fonts from the FontSize class and lazy-load them when needed to avoid memory spikes.
* Pie chart with horizontal serie direction did not use different colors when the VaryColors property was set.
* ExcelWorksheet.Dimension didn't correctly determine sheet dimension if data resides on last excel row.
* Hyperlinks referencing multiple cells was only loaded for the first cells in a range.
* Improve SUBTOTAL handling of cells hidden by filters.

## Version 5.8.5
### Minor Features.
* LoadFromCollection with attribute - Added EPPlusTableColumnSortOrder which enables column sort order on class level.
### Fixed issues
* ExcelDrawings.AddBarChart method for pivot charts had the wrong signature.
* Formula calculation adjusted to Excels behaviour regarding when 0 is returned instead of null. Thanks to Colby Hearn for the PR!
* The ExcelRange.Clear method did not remove threaded comments.
* Setting ExcelRange.Value to null only sets the value of the top-left cell. 
* Fixed an issue with pivot cache fields having an empty header in the source when updating the cache.
* Pivot field cache containing float and null values caused a corrupt workbook.
* EPPlus could not open workbooks without a normal style.
* SUMIF cannot handle single value, bug fix via Colby Hearn's PR. #570 - Invalid handling of numeric strings in COUNTIF, COUNTIFS and AVERAGEIF
* MINIFS and MAXIFS now return zero when there are no matches. 
* Added support for setting LicenseContext using a process level environment variable.

## Version 5.8.4
### Features
* 6 new functions:
	* NORMINV
	* NORM.INV
	* NORMSINV
	* NORM.S.INV
	* NORMDIST
	* NORM.DIST
### Fixed issues
* Default behaviour of SUM, AVERAGE, AVERAGEA, MEDIAN, LARGE, SMALL and PRODUCT functions did not handle cells with errors correctly
* Fixed load/save of .xlsm files not having an vbaproject.bin in the package.
* EPPlus threw an exception when handling extlist logic for spaceseparated data validations.
* IntParser (formula calc.) could not handle boolean values
* Copy of comments within the same worksheet caused an ArgumentException when loading the workbook again.
* Range.Copy of conditional formattings with multiple addresses did not work.
* Changed Uri reference handling to avoid relative references to the root.

## Version 5.8.3
### Fixed issues
* Inserting rows into a worksheet sometimes didn't update addresses on workbook defined names.
* Overlapping data validation addresses was validated on save which cause workbooks containing such not to be saved.
* Failed to copy cells when data validations or conditional formatting was set.
* COUNTIFS, AVERAGEIFS and SUMIFS fails with single cell ranges.
* Packages with VBA project with a component reference with encoded characters causes the saved package to become corrupt.
* It was not possible to specify legend entry properties for items from secondary y-axis.

## Version 5.8.2
### Fixed issues
* Range.Text returned the wrong value for format #.##0"*";(#.##0)"*" on negative values
* LoadFromCollection re-ordered the columns when no order was specified and and the item had more than 16 columns.
* Adding an Unchecked CheckBox Control Created Invalid XLSX File

## Version 5.8.1
### Fixed issues & minor features
 * Support for complex types in LoadFromCollection with attributes.
 * Fixed a bug where ExcelFunction.ArgToDecimal rethrow other error types as #VALUE.
 * Improved handling of decimals in Concat operations during calculation.
 * High-Low lines / Up-Down bars and Droplines were not loaded from an existing package.
 * Removed validation for negative values in conditional formatting priority, as negative values should be allowed

## Version 5.8.0
### Features
* ExcelWorksheetView.SplitPanes method added
* ExcelRangeBase Fill method added
	* FillNumber
	* FillDateTime
	* FillList
* New collection properties for Rows and Columns
	* ExcelWorksheet.Rows
	* ExcelWorksheet.Columns
	* ExcelRangeBase.EntireRow
	* ExcelRangeBase.EntireColum
* Support for formatting and deleting individual Chart Legend Entries.
* Range.Copy improvments. 
* Handle complex types in LoadFromCollection with attributes
* The ExcelPackage constructor and the Load/Save methods will now take a string path argument as well as a FileInfo.

### Fixed issues
* Renaming table to a name which contains the old name doesn't correctly update column references
* Range.Text did not handle empty formats like ;;; correctly.
* Range.Text - Strings within a format with percent, 0"%", incorrectly divides by 100
* VBA module names that started with a underscore (_) caused the ExcelWorksheet.Copy method to fail.
* Using range.RichText.Remove did not reflect the text to Range.Text.
* Adding a column to table with one column did not add the column.
* ExcelRangeBase.SaveToText did not add TextQualifiers around a formatted numeric value containing the column separator.
* Deleting cells sometimes didn't delete comments.
* Improved handling of ranges as return values from functions in the formula calculation.

## Version 5.7.5
### Fixed issues
* ExcelTable.CalculatedColumnFormula were not updated when inserting/deleting rows and columns.
* Copying a worksheet- or adding a VBA module- with a name starting with a number caused the macro-enabled workbook to be corrupt.
* LoadFromCollection with attributes didn't create a table if TableStyle was none.
* Fixed LEN functions incorrect handling of cached addresses.
* Fixed handling of worksheet names with escaped single quotes in the formula parser.
* ExcelPicture.SetSize sets an incorrect width when having a non-default dpi on an image.
* ExcelColumn.ColumnMax was not correctly set when deleting a column within a column span.
* Updating a pivot cache share among multiple pivot tables did not update/remove fields if the source range fields had been altered.
* Clearing or overwriting a cell in a shared formula range did not work correctly.
* Formulas in conditional formatting and data validations were not updated when inserting and deleting rows and columns.
* When deleting columns, defined names were not always correctly updated.
* ExpressionEvaluator could not handle a leading equal operator in combination with wildcards

## Version 5.7.4
### Fixed issues
* Metadata will now be removed from any formula copied to an external workbook to avoid corruption.
* Renaming table now updates references in formulas.
* ExcelNamedRange.Equals now works as expected.
* Inserting rows that shift conditional formatting outside the worksheet now adjust addresses correctly.
* Inserting rows into table formulas will not set the address to #REF!
* COUNTA function will now count errors and empty strings.
* LoadFromText and SaveToText did not handle quotes correctly.
* Changing the font of the current theme and the normal style does not always reflect when styling empty cells.
* Workbooks are getting corrupted when creating a pivot table where some cells have length greater than 255 characters.
* Deleting ranges with conditional formatting with multiple addresses sometimes threw an ArgumentOutOfRangeException.
* Copying a comment only add the text and the author, leaving any styling set on the comment.

## Version 5.7.3
### Fixed issues & minor features
* Add static methods to ExcelEncryption to encrypt and decrypt raw packages
	* ExcelEncryption.EncryptPackage
	* ExcelEncryption.DecryptPackage
* Conditional formatting lost some styles and added hair borders to empty border elements.
* COUNTBLANK and other functions using the ExpressionEvaluator don't handle time values correctly.
* COUNTBLANK does not handle cached addresses correctly.
* LoadFromText and SaveToTest did not work correctly with apostrophes.
* Changing the font of the normal style and then create a new named style did not inherit the font correctly.
* Data validations were shifted down instead of right when inserting cells into a range
* Comments and threaded comments were not shifted correctly to the right when inserting cells.
* EPPlus will now throw an exception if merging a range that overlaps a table.

## Version 5.7.2
### Fixed issues
* Pivot cache fields that contains both int's and float's corrupts the pivot cache.
* Added new methods to themes major- and minor- font collection - SetLatinFont, SetComplexFont, SetEastAsianFont, Remove and RemoveAt.
* Null or non existing external references to images on picture objects causes save to crash.
* VBA projects with the "dir" stream containing the unhandled value 0x4a, caused the workbook to become corrupt.
* Defined names with prefix and external reference throw a NullReference on load.

## Version 5.7.1
### Fixed issues
* Using a number format with AM/PM resulted in an output of AM or PM only.
* Validation of data validations throw an exception when Formula1 is empty even if AllowBlank was set to true
* Table behavior is incorrect when inserting rows if another table is below.
* Table calculated columns don't update the formula for added rows.
* Added cache for texts/fonts in AutoFitColumns, thanks to Simendsjo
* Setting Range.IsRichText on ranges with more than one cell did not work correctly.
* Loading packages with external references that didn't have a valid Uri failed on load.
* A value of null in a cell returned "0" in the Text property.

## Version 5.7.0
### Features
* External links
	* Adding, removing and break links to external workbooks.
	* Updating external workbook value/defined name caches.
	* Using external workbook caches in the formula parser.
	* Using loaded external packages (workbooks) in the formula parser.
* Enhanced sorting 
	* Pivot table auto sort - Sort on data fields using pivot areas.
	* Sort state in tables and auto filters. 
	* Left-to-right sorting and sorting with custom lists.

* Support for Show Data As on pivot table data fields
### Fixed issues
* Support for ErrorBars on both X and Y axis on scatter, bubble and area charts.
* Copying a worksheet to a new package fails with named styles in some cases. 
* When having multiple identical named styles (with the same styles set) only the first is copied.
* Handling of DBNull.Value in LoadFromDataTable function now works as expected.
* Added HideDropDown property to Data validation of type list.
* Added better exception messages for access to a package that has been disposed.
* Handling drawing objects had a concurrency when the size and position were asjusted.
* Pivot table caches with numeric/int and null values got corrupt due to an incorrect value on the containsMixedTypes attribute.

## Version 5.6.4
### Fixed issues
* Setting TableColumn.CalculatedColumnFormula doesn't set the formula correctly on the range.
* SUMIF function could not handle arrays or ranges of criteria.
* Fixed a bug in the Tokenizer where arrays of strings did not work properly.
* Comments were removed from cells when the IsRichText flag was set.
* AutoFitColumns did not calculate the font widths correctly for unknown font sizes.
* AutoFitColumns did not take DPI settings into account.
* Some styles did not get applied in Libre Office do to missing apply attributes on the style.
* Merged cells could set duplicate ranges causing the workbook to be corrupt.
* Vml documents containg unclosed &lt;BR&gt; tags failed to load.

## Version 5.6.3
### Fixed issues
* Changing the Normal style does not reflect correctly to cells with no style. 
* Formula calculation does not ignore nested SUBTOTAL values when more than one level.
* Reading a workbook with a pivot table slicer can cause the document to be corrupt on save
* Invalid handling in the RowHeightCache for resizing drawings caused KeyNotFoundException in some cases.
* Reverted Microsoft.Extensions.Configuration.Json for .NET standard 2.0 to 2.1.1 for support with ASP.NET Mvc Core 2.1
* The calculation DependencyChain caused an unhandled exception with workbook defined names and when having references to deleted worksheets.

## Version 5.6.2
### Fixed issues
* InvalidOperationException is thrown if Data Validation formula length exceeds 255 characters due to added _xlfn/_xlws prefixes. 
* Pivot tables with row/column fields of an enum type corrupted the pivot cache.
* Reading conditional formatting from a workbook causes fonts and number formats to clear
* Custom data validation could not handle formula reference to other worksheet.
* Normal build in style causes corruption when a style named Normal exists.
* Optimize handling of image sizing and positioning on save and resizing.
* Exception on copying worksheet with a chart that has no valid X Series.
* Match function could not handle named ranges in the range argument.
* Preserve table names when copying worksheet to a new workbook. 
* Number formats with localization did not set the culture code correctly.

## Version 5.6.1
### Fixed issues & minor features
* Styles xml could get corrupt in some cases as the numFmt dxf element was created in the wrong order according to the schema.
* Merged cells could get index out of range if deleting a merged area.
* LoadFromCollection will now set the cell's Hyperlink property for class members of type Uri or ExcelHyperlink.
* EPPlus will now preserve the 'aca' and 'ca' attributes for array formulas.
* ExcelRange.Formula and ExcelRange.FormulaR1C1 didn't return a value for array formulas except for the first cell in the range.
* Defined names containing #REF! throw an exception when copying a worksheet.
* Worksheet.FirstSheet property was set to the first sheet visible, if worksheet with position 0 was hidden.

## Version 5.6.0
### Features
* Custom table styles.
	* Create and modify names table styles that can be applied to tables and pivot tables.
	* Create and modify named slicer styles
* Pivot table styling using pivot areas.
* Enhanced table styling
* Added three new style properties to the ExcelTable and ExcelTableColumn
	* HeaderRowStyle
	* DataStyle 
	* TotalsRowStyle

### Fixed issues
* Using references to tables in formulas did not work correctly in some cases.
* Functions SUMIFS and AVERAGEIFS ignored hidden cells.
* SUMIFS could not handle horizontal ranges.
* SUMIFS-ExpressionEvaluator could not handle space between operator and criteria. Usage of the Dimension.End.Rows property led to premature stoppage in some cases
* INDEX function should round decimal numbers to integers using floor.
* Calculating shared formulas referencing full columns or full rows gives #ref! on all cells but the first 
* Loading worksheets from QXlsx failed.
* Deleting rows with comments throwed an exception in some cases.
* Deleting worksheets with comments throwed an exception in some cases.
* Removed validation for negative less that -60 in Drawing.SetPosition for RowOffset and ColumnOffset


## Version 5.5.5
### Fixed issues
* Dependency chain sometimes drops refererences when cross-worksheet addresses are used with defined names in the formula parser.
* ExcelWorksheetView ShowHeaders had an incorrect default value causing it not to work.
* Fixed issue when EPPlus crashes on load if a pivot table uses an external source.
* EPPlus will now preserve the 'cm' and 'vm' attributes of the sheet xml - 'c' element

## Version 5.5.4
### Fixed issues
* Formula calculation returns 0 in some cases when Excel returns null/empty in for example the MIN and MAX functions
* Fixed a bug in SUMIFS function that occured with more than 3 criterias
* The INDEX function could not handle that the range covered the cell containing the formula.
* AutoFilter.ShowButton property did not work.
* Number format "(#,##0)" got incorrectly formatted with a "-" prefix in the ExcelRange.Text property.
* Load using some CultureInfo did not work properly due to different StringComparation behavior.


## Version 5.5.3
### Fixed issues
* Min and Max could not handle empty ranges. They now returns 0 like Excel does if the range is empty.
* Added fallback for encoding of unknown unicode characters when saving shared strings. Thanks to SamSaint.
* ExcelTableColumnCollection.Insert used an invalid key when creating the name dictionary. Thanks Meigyoku-Thmn
* Fixed  ' object reference not set to an instance of an object ' if a middle column cell is null in the ToDataTable method. Thanks to Mengfw
* Added support for handling multiple colons in addresses, e.g. a1:a5:b3
* Fixed handling of the tilde(~) char in WildcardValueMatcher when not being an escape character. Used by If functions.
* Setting styles over empty columns caused cell to be removed in some cases.
* Having pivot table shared items with both null and empty string causes an exception on load.
* Defined names referencing non-existing worksheets caused crash on load.
* Exposed static method RecyclableMemory.SetRecyclableMemoryStreamManager. Thanks to LIFEfreedom.

## Version 5.5.2
### Fixed issues
* Fixed a bug in ExpressionEvaluator that caused search criteria's starting with a wildcard character to cause an Exception.
* Setting cell value via an array overwriting rich text causes invalid cell content.
* Removed invalid handling of defined names on load if the name contained a link to an external workbook.
* Dependency chain referenced the wrong worksheet in some cases when a formula had off-sheet references when checking for circular references.
* Table headers with rich text caused corrupt workbook.
* Fixed error handling and handling of single cell addresses in COUNTIFS function
* Downgraded .NET 5 referenced packages for .NET standard 2.0 and 2.1 build.

## Version 5.5.1
### Features
* 10 new functions:
	* COMPLEX
	* DEVSQ
	* AVEDEV 
	* GAMMALN 
	* GAMMALN.PRECISE
	* GAMMA
	* DB 
	* SHEET
	* INTRATE
	* MDURATION
### Fixed issues
* LoadFromCollection thrown an exception if collection was empty and PrintHeaders was false
* Adding controls after an table was added caused an corrupt workbook.
* PrecisionAndRoundingStrategy.Excel doesn't work in range -1 to 1
* EPPlus no longer validate defined names on load.
* Setting the IsRichText property to true don't convert the value to Rich text properly.
* Xml formatting issue on Saving DrawingXml. Thanks to ZaL133 for pull request.

## Version 5.5.0
### Features
* Form Controls
	* Buttons
	* Drop-Downs
	* List Boxes
	* Check Boxes
	* Radio Buttons
	* Spin Buttons
	* Scroll Bars
	* Labels
	* Group Boxes 
* Group/Ungroup drawing object via the ExcelDrawing.Group and ExcelDrawing.Ungroup methods
* New attribute support for LoadFromCollection. See https://github.com/EPPlusSoftware/EPPlus/wiki/LoadFromCollection-using-Attributes
* 20 new functions 
	* AGGREGATE
	* PERCENTILE.EXC
	* DATEDIF
	* QUARTILE.EXC
	* STDEVA
	* STDEVPA
	* VARA 
	* VARPA
	* COVAR 
	* COVARIANCE.P
	* COVARIANCE.S
	* RANK
	* DOLLAR
	* DOLLARDE
	* DOLLARFR
	* PERMUT 
	* PERMUTATIONA
	* XOR
	* MULTINOMIAL
	* YIELDMAT

### Fixed issues & minor fixes
* Round methods can now use only 15 sigificant figures with calculation option - PrecisionAndRoundingStrategy  
* R1C1 causes corrupt worksheet if used as shared formula with off-sheet reference. 
* Using double qoutes in R1C1 didn't encode correctly.
* Altering fields on a table connected to a pivot table causes the pivot table to become corrupt.
* Pivot tables with a boolean column and a filter got corrupt on save. 
* Deleting a worksheet and adding it again with the same name causes a crash on save if the worksheet is referenced by a pivot table. This happends due to the SourceRange property still referencing the deleted worksheet.
* Changed ExcelAddressBase.FullName to public
* Reference table name only in indirect formula did not work.
* Replaced MemoryStrem internally with Microsoft.IO.RecyclableMemoryStream. Thanks to LIFEfreedom.
* Referencing a single cell with rich text from a formula returned an invalid value.
* Reverted .NET 5.0 references for .NET standard 2.0 and 2.1, to be compatible with Azure Functions.

## Version 5.4.2
### Fixed issues
* .NET 5.0 to Nuget package.
* Deselecting pivot table items did not work for null values in some cases
* Deleting a worksheet containing a pivot table was not properly cleaned.
* Save crashed if a pivot tables shared items had the same string value with different case.
* Reading Hyperlinks in the format #'sheet'!A1 will now work.
* Fixed an issue with some table properties overwiting other table properties
* Fixed an issue with Range.SaveToText hanging with the FileInfo overload
* ExcelShape.SetFromFont did not set fonts correctly
* Inserting more than 32K rows at once did not work.
* Worksheets with gaps of more than 32K rows causes invalid values to be returned in some cases.

## Version 5.4.1
### Minor new features
* WriteProtection added to Workbook.Protection. Allows to set a workbook to read-only with a password.
* ToDataTable method added to ExcelRange
### Fixed issues
* Worksheet names starting with R1C1 format creates invalid addresses
* Row array in pages in the cell store caused an index out of bounds exception in rare cases.
* Setting a shared formula with an external reference causes the workbook to become corrupt in some Excel versions 
* A Pivot table get corrupt if a TimeSpan is used in a column that needs shared items in the cache.
* Cell.Text returns an incorrect text when formatted #,##0;#,##0;- and value is rounded to zero.
* Delete Table With ExcelTableCollection.Delete() did not remove the xml in the package causing the table not to be deleted.
* Pivot tables with 'Save source data with file' caused an exception.
* Signing VbaProject under .NET core did not work correcly.

## Version 5.4.0
### Features
* Pivot tables filters
	* Item filters - Filters on individual items in row/column or page fields.
	* Caption filters (label filters) - Filters for text on row and column fields.
	* Date, numeric and string filters - Filters using various operators such as Equals, NotBetween, GreaterThan, etc.
	* Top 10 filters - Filters using top or bottom criterias for percent, count or value.
	* Dynamic filters - Filters using various date- and average criterias.
* Add calculated fields to pivot tables.
* Support for pivot table shared caches via the new overload for PivotTables.Add.
* Support for pivot table slicers
* Support for table slicers
### Minor new features
* Number format property has been added to pivot fields and pivot cache fields.
* Added PivotTableStyle property and enum for easier setting all pivot table styles

### Fixed issues
* Chart series indexes got corrupt when removing and adding series in some cases.
* Inheriting cells styling from column to cell did not work correctly in some cases.

## Version 5.3.2
### Fixed issues
* Workbook got corrupt on copy if a worksheet VBA Code module was null.
* Setting chart series color settings to WithinLinear caused an exception in some cases.
* LoadFromCollection did not load all members as in previous version when suppling binding flags and MemberInfo's.
* If first worksheet(s) are hidden in the worksheet collection, print preview crashed in Excel.
* Worksheet Move methods failed with a crash in some cases.

## Version 5.3.1
### Features
* Added support for copying threaded comments 
* Add FormulaParserManager.Parse method.
### Fixed issues
* Formulas and formula values did not encode characters below 0x1F correctly
* ExcelPackage.LoadAsync did not work with non seekable streams
* A chart that reference within its own worksheet will now change the worksheet in the series addresses for any copy made with the Worksheets.Add method
* ExcelRange.ToText method did not work correctly with rich text cells.

## Version 5.3.0
### Features
* Support for threaded comments with mentions. See https://github.com/EPPlusSoftware/EPPlus/wiki/Threaded-comments
* Support for two new functions:
    * MINIFS
    * MAXIFS
### Fixed issues
* Removed the extra comman added to the end of the header row in the ExcelRange.ToText method.
* The abstract class ExcelChart missed the properties DisplayBlankAs, RoundedCorners, ShowDataLabelsOverMaximum, ShowHiddenData after the version 5.2 update. The properties have been moved back again.
* ExcelColorXml.Indexed returned 0 if was not set, causing the LookupColor method to return an incorrect color.

## Version 5.2.1
### Features
* New method ExcelRange.LoadFromDictionary. 
* Support for Dynamics in ExcelRange.LoadFromCollection (ExpandoObject).
* New Lambda config parameter to ExcelRange.LoadFromCollection with new options for header parsing (for example: UnderscoreToSpace, CamelCaseToSpace)
* Zip64 support in packaging to enable larger packages.
### Fixed issues
* Wildcards in MATCH function not working
* Performance issue in ExpressionEvaluator.GetNonAlphanumericStar
* Using the sheet name to reference the entire worksheet did not work in formulas.
* GetAsByteArray corrupted the package if ExcelPackage.Save was called before.
* Parsing issue in the Value function

## Version 5.2.0
### Features
* Support for extended chart types and Stock charts:
	* Four types of stock charts: High-Low-Close, Open-High-Low-Close, Volume-High-Low-Close and Volume-Open-High-Low-Close
	* Sunburst Chart
	* Treemap Chart
	* Histogram Chart
	* Pareto Chart
	* Funnel Chart
	* Waterfall Chart	
	* Box &amp; Whisker Chart
	* Region Map Chart

* Support for 49 new functions:
    * BESSELI
    * BESSELJ
    * BESSELK
    * BESSELY
    * COUPDAYBS
    * COUPDAYS
    * COUPDAYSNC
    * COUPNCD
    * COUPNUM
    * COUPPCD
    * CUMIPMT
    * CUMPRINC
    * DDB
    * DISC
    * DURATION
    * EFFECT
    * ERF
    * ERF.PRECISE
    * ERFC
    * ERFC.PRECISE
    * FV
    * FVSCHEDULE
    * IPMT
    * IRR
    * ISPMT
    * MIRR
    * MODE
    * MODE.SNGL
    * NOMINAL
    * NPER
    * NPV
    * PDURATION
    * PERCENTILE
    * PERCENTILE.INC
    * PERCENTRANK
    * PERCENTRANK.INC
    * PPMT
    * PRICE
    * PV
    * RATE
    * RRI
    * SLN
    * SYD
    * TEXTJOIN
    * VAR.P
    * VAR.S
    * XIRR
    * XNPV
    * YIELD

* Four new properties to Style.Font
	* Charset
	* Condense
	* Extend
	* Shadow
* Added   * New `As` property to the Drawing, ConditionalFormatting and DataValidation objects for easier type cast. For example `var sunburstChart = worksheet.Drawings[0].As.Chart.SunburstChart;` or `var listDv = sheet.DataValidations.First().As.ListValidation;`
* OFFSET function calculation can now be a part of a range in formula calculation. For example `SUM(A1:OFFSET(A3, -2, 0))`
* Added ShowColumnHeaders, ShowColumnStripes, ShowRowHeaders, ShowRowStripes and ShowLastColumn properties to ExcelPivotTable
### Fixed issues
* Ignore a leading "_xlfn." in function names when calculating formulas.
* Only the fist paragraph was loaded when the RichText property was loaded from the XML.
* VerticalAlign property did not delete the XmlNode properly when set to None.
* A named style with a missing fontid crashed on load.
* Deleting a worksheet in a macro enabled workbook causes a NullReferenceException.

## Version 5.1.2
### Features
* Added ClearFormulas and ClearFormulaValues to Range, Worksheet and Workbook. ClearFormulas removes all formulas, ClearFormulaValues removes all previously calculated/cached values.
* Added support for 19 new engineering functions: 
	* CONVERT
	* DELTA
	* BIN2DEC
	* BIN2HEX
	* BIN2OCT
	* DEC2BIN
	* DEC2HEX
	* DEC2OCT
	* HEX2BIN
	* HEX2DEC
	* HEX2OCT
	* OCT2BIN
	* OCT2DEC
	* OCT2HEX
	* BITLSHIFT
	* BITAND
	* BITOR
	* BITRSHIFT
	* BITXOR
### Fixed issues
* Delete method adjusted row offset on drawings incorrectly.
* When copying a worksheet with images to an other package did not work correctly in some cases.
* Datavalidation addresses did not update correctly when deleting columns.
* Opening a packages saved with System.IO.Compression caused a corrupted package on save. 
* Added a check to the ExcelPackage Construcors if the FileInfo exists and is zero-byte. Supplying a zero-byte file will now create a new package. Supplying a zero-byte template will throw an exception.
* Fixed scaling for pictures. Changed data type for internal pixel variables from int to double.
* Delete/Insert din't handle comma separated addresses in data validation / conditional formatting
* ColumnMin and ColumnMax were not correctly updated when one or more columns were deleted.

## Version 5.1.1
### Features
* Added method RemoveVBAProject to ExcelWorkbook.
### Fixed issues
* CalculatedColumnFormula property was set to the range on save, overwriting any cell value that was changed in the range of the formula.
* VbaProject. Remove didn't fully remove the VBA project.
* LoadFromCollection didn't work will List&lt;object&gt;.
* Group shapes containg shapes with the same name throw exception.
* Worksheets with a large amount of columns had bad performance on save.

## Version 5.1.0
### Features
* Insert and Delete method added to ExcelRange. Cells will be shifted depending on the argument supplied.
* AddRow, InsertRow and DeleteRow added to ExcelTable.
* Add, Insert and Delete added To ExcelTableColumnCollection.	
* Added support for new functions: 
	* FACTDOUBLE
	* COMBIN
	* COMBINA
	* SEC
	* SECH
	* CSC
	* CSCH
	* COT
	* COTH
	* RADIANS
	* ACOT
	* ACOTH
	* ROMAN
	* GCD
	* LCM
	* FLOOR.PRECISE
	* CEILING.PRECISE
	* MROUND
	* ISO.CEILING
	* FLOOR.MATH
	* CEILING.MATH
	* SUMXMY2
	* SUMX2MY2 
	* SUMX2PY2
	* SERIESSUM

### Fixed issues
* Images added with AddImage(Image) did not use the oneCellAnchore element.
* ExcelPackage.CopyStreamAsync did not fully use async (Flush-->FlushAsync), causing an exception.
* VBA module names restricts some allowed characters.
* Shared Formulas are not handled correctly when an address argument reference another worksheet.
* Adding a Sparklinegroup does not add it to the SparklineGroups collection.
* Files saved in LibreOffice did not handle boolean properties correctly for rows and columns, (for example the hidden property).
* Data validation - List validation did not set the ShowErrorMessage when an address referenced another worksheet.
* Addresses with style: 'sheet'!A1:'sheet'!A2 was not handled correctly.

## Version 5.0.4
### Fixed issues
* Datavalidation on lists failed if the formula was an defined name.
* Merged cells got cleared if a value was set over multiple cells
* RichText causes xml "corruption" if a blank string or null was added to the collection.

## Version 5.0.3
### Minor new features
* Chart series will from version 5 handles both addresses and arrays. Arrays are handled in the StringLiteralsX, NumberLiteralsX and NumberLiteralsY arrays when a series is set to an Array ( for example {1,2,3} ).
* Add support for BringToFront and SendToBack methods on the Drawings Collection to handle drawing overlap.
* Add TopLeftCell to ExcelWorksheetView
* Enabled Style.TextRotation=255 for Vertical Text in cells. Added new method ExcelStyle.SetTextVertical.
* Add Pivot property to conditional formatting, to indicate a pivot source.
* RichText on drawings can now handle paragraphs to get line breaks. Add method has new parameter NewParagraph.
* Table source overload to PivotTable.Add method
* 13 new functions supported in the Formula calculation engine: IFS, SWITCH, WORKDAY.INTL, TYPE, ODD, EVEN, DAYS, NUMBERVALUE, UNICHAR, UNICODE, CLEAN, TRIM and CONCAT
### General fixes
* Cellstore has been rewritten. This should fix some issues with inserting and deleting rows and columns. Also fixes a sorting issue.
* Fixed SchemaNodeOrder in many drawing classes.
* Handling of [circular references](https://github.com/EPPlusSoftware/EPPlus/wiki/Circular-references) has been redesigned.
### Fixed issues
* Worksheet.Hidden does not always hide the worksheet.
* Drawings will now move and size when inserting/deleting rows/columns depending on the ExcelDrawing.EditAs property.
* Adding a Table caused an exception if a chart sheet existed in the workbook.
* Adding a PivotTable caused an exception if a chart sheet existed in the workbook.
* Custom document properties are case insensitive.
* Sheet with rich text inline string can not handle whitespaces only.
* StackOverflowException when calling Clear on a comma separated Range.
* ExcelWorksheet.Copy corrupts package if a relationship to drawing.xml with no drawings exists.
* Copying formulas in ranges can lead to invalid #REF! for fixed addresses.
* Table column names are validated for duplicates on a non-encoded value
* Packages with a prefix for the main xml namespace for a worksheet gets corrupted.
* ExcelRangeBase.AutofitColumns() unhides hidden columns
* The Normal style does not work correctly if not named Normal. The Normal style is now found using the first occurrance of the BuildInId = 0 criteria.
* URI for the sharedstrings and the styles part were not fetched by RelationshipType when a package was loaded.
* Setting a cell value to a char datatype will result i "0" when saved
* Structs in a cell value can result in a null value when converted to string on save
* Conditional formatting styles crashed when copying a worksheet from another package.
* EPPlus crashes on load if a workbook or worksheet has more than one defined name with the same name.
* Row styles were not copied correctly copied when inserting rows
* Overwriting a shared formulas first cell causes a crash.
* Workbooks with Empty series for Scatter- and Doughnut- charts crashes on load
* FileStream for compound documents are not closed.
* If CustomSheetView element contained row/column breaks, the package could not be loaded.
* Pivot tables crashed if SubTotalFunction were set to eSubTotalFunctions.None and there were no values in the source data.
* Data validations of type list don't support formula references to other worksheets.
* _Formula calc_: Does not remove double-negation from formulatokens.
* _Formula calc_: Value matcher now supports comparisons between DateTime and double. CompileResultFactory includes float type in DataType.Decimal.
* _Formula calc_: MultipleRangeCriterasFunction.GetMatchIndexes() looped through max number of rows when a range argument was an entire column now stops as Dimension.End.Row. Fixed a bug in CountIfs function which wasn't working properly with multiple criteria's
* _Formula calc_: Support Instance_num parameter of SUBSTITUTE function.