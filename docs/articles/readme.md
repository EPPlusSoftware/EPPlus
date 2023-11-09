# EPPlus 7

## Announcement: new license model from version 5
EPPlus has from this new major version changed license from LGPL to [Polyform Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0/).

With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.

This is explained in more detail [here](https://www.epplussoftware.com/Home/LgplToPolyform).

Commercial licenses, which includes support, can be purchased at (https://www.epplussoftware.com/).

The source code of EPPlus can be found at our [github repository](https://github.com/EPPlusSoftware/EPPlus)

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
This might be the easiest way of configuring this. Just as above, set the variable to Commerical or NonCommercial depending on your usage.

**Important!** The environment variable should be set at the user level.


## Breaking changes EPPlus 7
See [Breaking Changes in EPPlus 7](https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-7)

## New features in EPPlus 7
EPPlus 7 comes with a set of new features, see (https://www.epplussoftware.com/Developers/Features)

## Improved documentation
EPPlus 7 has new, separate sample projects for [C#](https://github.com/EPPlusSoftware/EPPlus.Samples.CSharp) and [Visual Basic](https://github.com/EPPlusSoftware/EPPlus.Samples.VB) respectively.
There is also an updated [developer wiki](https://github.com/EPPlusSoftware/EPPlus/wiki). The work with improving the documentation will continue, feedback is highly appreciated!

