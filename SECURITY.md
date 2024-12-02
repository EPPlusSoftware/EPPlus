# Security Policy

_Last updated: 2024-08-14_

## Supported Versions

EPPlus 5, 6 and 7 are automatically scanned for vulnerabilities and static code analysis is performed as part of the CI. 

| Version | Supported          | Comment            | Deprecation date |
| ------- | ------------------ | ------------------ |----|
| 7.x.x   | :white_check_mark: |                    ||
| 6.x.x   | :white_check_mark: |                    |2025-12-31|
| 5.x.x   | :white_check_mark: |                    |2024-12-31|
| < 4.3   | :x:                |Deprecated/unsupported versions|2020-12-31|

## Security update policy
Security patches will be provided via new revisions released in our public Nuget feed. One patch for each supported major version/the two latest minor versions will be provided. [Subscribe to our newsletter](https://epplussoftware.com/en/Home/Newsletter) to get updates from EPPlus Software.

## Reporting a Vulnerability

Create an issue in our [issue tracker](https://github.com/EPPlusSoftware/EPPlus/issues), describe the vulnerability (including relevant links) and what version of EPPlus that is affected.

## Code signing 
Since version 7.5 the EPPlus Nuget package and the EPPlus libraries/dll:s are digitally signed by EPPlus Software AB.

## See also
- [EPPlus versioning](https://github.com/EPPlusSoftware/EPPlus/wiki/Releases-versioning)

## Vulnerabilities
|Detected|Resolved|Affected EPPlus versions|CVE|Our comment|Resolution|
|--------|--------| ----------------------|---|----------|----------|
|October 10, 2024|October 11, 2024|EPPlus 7.x,targeting .NET 7 or 8|[Microsoft Security Advisory CVE-2024-38095](https://github.com/advisories/GHSA-447r-wph3-92pm) and [Microsoft Security Advisory CVE-2024-30105](https://github.com/advisories/GHSA-hh2w-p6rv-4g7w)|Microsoft has released a security fix in Microsoft.Extensions.Configuration.Json 8.0.1. The potential risk for most users should be low.|Patch  released in version 7.4.1|
|September 9, 2024||EPPlus 7.x, targeting .NET 7 or 8|[Microsoft Security Advisory CVE-2024-38095](https://github.com/advisories/GHSA-447r-wph3-92pm) and [Microsoft Security Advisory CVE-2024-30105](https://github.com/advisories/GHSA-hh2w-p6rv-4g7w)|Microsoft has released security fixes for System.Text.Json and System.Formats.Asn1 (transient dependencies in EPPlus). The potential risk for most users should be low.|Patch  released in version 7.3.2|
|June 15, 2023|June 15, 2023|EPPlus 6.x prior to 6.2.6, targeting .NET 6 or 7.|[.NET Denial of Service vulnerability (CVE 2023-29331)](https://github.com/advisories/GHSA-555c-2p6r-68mm)|Microsoft has released a security fix for a Denial of Service vulnerability (CVE-2023-29331) in System.Security.Cryptography.Pkcs for .NET 6 and .NET 7. EPPlus uses this component for x509 certificates used when signing VBA projects in a workbook. The potential risk for most users should be low, as the certificates used to sign your workbooks are usually known.|Upgrade to EPPlus 6.2.6 or higher|
