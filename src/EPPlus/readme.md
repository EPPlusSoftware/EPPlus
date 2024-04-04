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

## New features in EPPlus 7
	* Calculation engine update to support array formulas. https://epplussoftware.com/en/Developers/EPPlus7
		* Support for calculating legacy / dynamic array formulas.
		* Support for intersect operator.
		* Support for implicit intersection.
		* Support for array parameters in functions.
		* Better support for using the colon operator with functions.
		* Better handling of circular references
		* 90 new functions
		* Faster optimized calculation engine with configurable expression caching.
		* Breaking changes: Updated calculation engine, See [Breaking Changes in EPPlus 7](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-7) for more information.
		* Conditional Formatting improvements
		* Improved performance, xml is now read and written on load and save.
		* Cross worksheet support formula support.
		* Extended styling options for color scales, data bars and icon sets.

## Breaking Changes
See https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-7

## Improved documentation
EPPlus 6 has a new web sample site available here: (https://samples.epplussoftware.com/) ,  Source code is available here: [EPPlus.WebSamples](https://github.com/EPPlusSoftware/EPPlus.WebSamples)
There is also a new sample project for four different docker images, [EPPlus.DockerSample](https://github.com/EPPlusSoftware/EPPlus.DockerSample)
EPPlus also has two separate sample projects for [.NET Core](https://github.com/EPPlusSoftware/EPPlus.Sample.NetCore/tree/version/EPPlus6.0) and [.NET Framework](https://github.com/EPPlusSoftware/EPPlus.Sample.NetFramework/tree/version/EPPlus6.0) respectively.
There is also an updated [developer wiki](https://github.com/EPPlusSoftware/EPPlus/wiki). 
The work with improving the documentation will continue, feedback is highly appreciated!