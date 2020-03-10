# Minor features/Fixed issues - EPPlus 5.0

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

## Version 5.0.4
### Fixed issues
* Datavalidation on lists failed if the formula was an defined name.
* Merged cells got cleared if a value was set over multiple cells
* RichText causes xml "corruption" if a blank string or null was added to the collection.