# EPPlus 8

## License
EPPlus 8 has a dual license model with a community license for noncommercial use: [Polyform Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0/).

With this license EPPlus is free to use for personal or noncommercial use, but will require a commercial license to be used in a commercial business.

Commercial licenses, which includes support, can be purchased at (https://www.epplussoftware.com/).

The source code for EPPlus is available at [EPPlus Software's github repository](https://github.com/EPPlusSoftware/EPPlus)

## License parameter must be set
Before using EPPlus 8, you must specify the license to use. This is done via the License property of the ExcelPackage class

For commercial use, you use the License.SetCommercial(string), with your license key as argument. 
Your license key is available on your license, under the section "My Licenses" on our website.

For noncommercial use, you set the License.SetNonCommercialOrganization(string) or License.SetNonCommercialPersonal(string) with the name as argument. 
Noncommercial use will reserve the Comment and Tag field of the package for license information and add a license file within the package.

You can also configure these settings in the configuration files or in an environment varialble:

### 1. Via code
```csharp
// If you are a commercial business and have
// purchased commercial licenses use the static property
// LicenseContext of the ExcelPackage class :
ExcelPackage.License.SetCommercial("<Your License Key here>");

// If you use EPPlus in a noncommercial context
// according to the Polyform Noncommercial license:
ExcelPackage.License.SetNonCommercialPersonal("<Your Name>");
//or..
ExcelPackage.License.SetNonCommercialOrganization("<Your Noncommercial Organization>");
    
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
            "License": "Commercial:<Your License Key here>" //The license context used
            }
        }
    }
}
```
### 3. Via app/web.config
```xml
<appSettings>
    <!--The license context used-->
    <add key="EPPlus:ExcelPackage.License" value="NonCommercialPersonal:Your Name" /> //..or use "NonCommercialOrganization:Your Organizations name" 
</appSettings>
```
### 4. Set the environment variable 'EPPlusLicenseContext'
This might be the easiest way of configuring this. Just as above, set the variable EPPlusLicense.

## New features in EPPlus 8
* Support for OLE objects (Linked or Embedded files).
* Support for digital signing workbooks and signature lines.
* In-cell pictures / support for the IMAGE function.
* Sensitivity Label API to integrate with MIP (Microsoft Information Protection SDK).
* Many minor features and bug fixes.
* 
## Breaking Changes
See https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-8

