# EPPlus 7

## Announcement: new license model from version 5
EPPlus has from this new major version changed license from LGPL to [Polyform Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0/).

With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.

This is explained in more detail [here](https://www.epplussoftware.com/Home/LgplToPolyform).

Commercial licenses, which includes support, can be purchased at (https://www.epplussoftware.com/).

The source code of EPPlus has moved to a [new github repository](https://github.com/EPPlusSoftware/EPPlus)

## LicenseContext parameter must be set
With the license change EPPlus has a new parameter that needs to be configured. If the LicenseContext is not set, EPPlus will throw a LicenseException (only in debug mode).

This is a simple configuration that can be set in a few alternative ways:

### 1. Via code
```csharp
// If you are a commercial business and have
// purchased commercial licenses use the static property
// LicenseContext of the ExcelPackage class :
ExcelPackage.LicenseContext = LicenseContext.Commercial;

// If you use EPPlus in a noncommercial context
// according to the Polyform Noncommercial license:
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    
using(var package = new ExcelPackage(new FileInfo("MyWorkbook.xlsx")))
{

}
```
### 2. Via appSettings.json
```json
{
    {
    "EPPlus": {
        "ExcelPackage": {
            "LicenseContext": "Commercial" //The license context used
            }
        }
    }
}
```
### 3. Via app/web.config
```xml
<appSettings>
    <!--The license context used-->
    <add key="EPPlus:ExcelPackage.LicenseContext" value="NonCommercial" />
</appSettings>
```
### 4. Set the environment variable 'EPPlusLicenseContext'
This might be the easiest way of configuring this. Just as above, set the variable to Commercial or NonCommercial depending on your usage.

**Important!** The environment variable should be set at the user or process level.

## New features in EPPlus 7 preview 1
* Calculation engine update to array formulas. https://github.com/EPPlusSoftware/EPPlus/wiki/EPPlus-7-Preview
	* Support for calculating legacy / dynamic array formulas.
	* Support for intersect operator.
	* Support for Implicit intersection.
	* Support for array parameters in functions.
	* Better support for using the colon operator with functions.
	* 21 new functions

## Breaking Changes
The formula parser has changed significantly in EPPlus 7, requiring all custom functions that inherits from the `ExcelFunction` class to be reviewed. 
The `ExcelFunction` class now exposes new properties used to handle array results and condition behaviour. 
* `NamespacePrefix` - If the function requires a prefix when saved, for example "_xlfn." or "_xlfn._xlws."
* `HasNormalArguments` a boolean indicating if the formula only has normal arguments. If false, the `GetParameterInfo` method must be implemented. Default is true.
* `ReturnsReference` - If true the function can return a reference to a range. Use the `CreateAddressResult` to return the result with a reference. Returning a reference, will cause the dependency chain to check the address and will allow the colon operator to be used with the function.
* `IsVolatile` -  If the function returns different result when called with the same parameters. Default false.
* `ArrayBehaviour` - If the function allows arrays as input in a parameter, resulting in an array output. Also see the `GetArrayBehaviourConfig` method.
Methods
* `CreateAddressResult`  - Return the result with an range to a range.
* `CreateDynamicArrayResult` - The result should be treated as a dynamic array.
* `GetArrayBehaviourConfig` - Sets the index if the parameters that can be arrays. Also see the `ArrayBehaviour` property.

* The source code tokenizer now tokenize more detailed, tokenizing addresses. 
* The expression handling is totally rewritten and now uses reversed polish notation instead of group expressions.


### Breaking Changes in version 6.
* All public references to System.Drawing.Common has been removed from EPPlus. See [Breaking Changes in EPPlus 6](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-6).
* Static class 'FontSize' has splitted width and heights into two dictionaries. FontSizes are lazy-loaded when needed. 
* ...and more, see https://epplussoftware.com/docs/6.0/articles/breakingchanges.html
### Breaking Changes in version 5.
* The default behavior for the Worksheet collection base in .NET Framework has changed from 1 to 0. This is the same default behavior as in .NET core today.
* Pictures have changed the behavior as the oneCellAnchor tag is used instead of the twoCellAnchor tag with the editAs="oneCell". 

## Improved documentation
EPPlus 6 has a new web sample site available here: (https://samples.epplussoftware.com/) ,  Source code is available here: [EPPlus.WebSamples](https://github.com/EPPlusSoftware/EPPlus.WebSamples)
There is also a new sample project for four different docker images, [EPPlus.DockerSample](https://github.com/EPPlusSoftware/EPPlus.DockerSample)
EPPlus also has two separate sample projects for [.NET Core](https://github.com/EPPlusSoftware/EPPlus.Sample.NetCore/tree/version/EPPlus6.0) and [.NET Framework](https://github.com/EPPlusSoftware/EPPlus.Sample.NetFramework/tree/version/EPPlus6.0) respectively.
There is also an updated [developer wiki](https://github.com/EPPlusSoftware/EPPlus/wiki). 
The work with improving the documentation will continue, feedback is highly appreciated!