# EPPlus 8

## License
EPPlus 8 has a dual license model with a community license for noncommercial use: [Polyform Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0/).

With this license EPPlus is free to use for personal/noncommercial use, but will require a commercial license to be used in a commercial business.

Commercial licenses, which includes support, can be purchased at (https://www.epplussoftware.com/). For more details on our licensing, see our [License FAQ](https://epplussoftware.com/en/LicenseOverview/LicenseFAQ).

The source code for EPPlus is available at [EPPlus Software's github repository](https://github.com/EPPlusSoftware/EPPlus)

## License parameter must be set
Before using EPPlus 8, you must specify the license to use. This is done via the static License property of the ExcelPackage class

For commercial use, set the ExcelPackage.License.SetCommercial(string), with your license key as argument. 
Your license key is available on your license, under the section "My Licenses" on our website and on EPPLus license documents created after February 2025.

For noncommercial use, set the ExcelPackage.License.SetNonCommercialOrganization(string) or ExcelPackage.License.SetNonCommercialPersonal(string) with your name as argument. 
When configured for noncommercial use, EPPlus will reserve the Comment and Tag field of the package for license information and add a license file within the package.

You can also configure these settings in the configuration files or in an environment varialble:

### 1. Via code
```csharp
// If you are a commercial business and have purchased a commercial license:
ExcelPackage.License.SetCommercial("<Your License Key here>");

// If you use EPPlus in a noncommercial context according to the Polyform Noncommercial license:
ExcelPackage.License.SetNonCommercialPersonal("<Your Name>");
// or...
ExcelPackage.License.SetNonCommercialOrganization("<Your Noncommercial Organization>");
    
using(var package = new ExcelPackage(new FileInfo("MyWorkbook.xlsx")))
{

}
```
### 2. Via appSettings.json
For commercial use:
```jsonc
{
    "EPPlus": {
        "ExcelPackage": {
            "License": "Commercial:<Your License Key here>" 
            }
        }
}
```
For noncommercial use:
```jsonc
{
    "EPPlus": {
        "ExcelPackage": {
            "License": "NonCommercialPersonal:<Your Name>" //..or use "NonCommercialOrganization:<Your organizations name>" 
            }
        }
}
```

### 3. Via app/web.config
For commercial use:
```xml
<appSettings>
    <!--The license context used-->
    <add key="EPPlus:ExcelPackage.License" value="Commercial:<Your License Key here>" /> 
</appSettings>
```
For noncommercial use:
```xml
<appSettings>
    <!--The license context used-->
    <add key="EPPlus:ExcelPackage.License" value="NonCommercialPersonal:<Your Name>" /> <!--..or use "NonCommercialOrganization:Your organizations name" -->
</appSettings>
```
### 4. Set the environment variable 'EPPlusLicense'
This might be the easiest way of configuring this.   
Set the variable to `Commercial:<Your License Key>`, `NonCommercialPersonal:<Your Name>` or `NonCommercialOrganization:<Your organizations name>`

The example below set the license key using the SETX command in the Windows console on the user level, but you can set it in any way you prefer.

```console
> SETX EPPlusLicense "Commercial:<Your License Key>"
```

## New features in EPPlus 8
See https://epplussoftware.com/en/Developers/EPPlus8
## Breaking Changes
See https://github.com/EPPlusSoftware/EPPlus/wiki/Breaking-Changes-in-EPPlus-8

