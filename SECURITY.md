# Security Policy

## Supported Versions

EPPlus 5 and 6 are automatically scanned for vulnerabilities and static code analysis is performed as part of the CI. Security patches will be provided via new revisions released in our public Nuget feed.

| Version | Supported          | Comment            |
| ------- | ------------------ | ------------------ |
| 6.x.x   | :white_check_mark: |                    |
| 5.x.x   | :white_check_mark: |                    |
| < 4.0   | :x:                |Deprecated/unsupported versions|

## Reporting a Vulnerability

Create an issue in our [issue tracker](https://github.com/EPPlusSoftware/EPPlus/issues), describe the vulnerability (including relevant links) and what version of EPPlus that is affected.

## Vulnerabilities
|Detected|Resolved|Affected EPPlus versions|CVE|Our comment|Resolution|
|--------|--------| ----------------------|---|----------|----------|
|June 15, 2023|June 15, 2023|EPPlus 6.x prior to 6.2.6|[.NET Denial of Service vulnerability (CVE 2023-29331)](https://github.com/advisories/GHSA-555c-2p6r-68mm)|Microsoft has released a security fix for a Denial of Service vulnerability (CVE-2023-29331) in System.Security.Cryptography.Pkcs for .NET 6 and .NET 7. EPPlus uses this component for x509 certificates used when signing VBA projects in a workbook. The potential risk for most users should be low, as the certificates used to sign your workbooks are usually known.|Upgrade to EPPlus 6.2.6 or higher|
