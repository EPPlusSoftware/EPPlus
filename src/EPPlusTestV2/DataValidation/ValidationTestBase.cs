/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Contracts;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace EPPlusTest.DataValidation
{
    public abstract class ValidationTestBase : TestBase
    {
        protected ExcelPackage _package;
        protected ExcelWorksheet _sheet;
        protected XmlNode _dataValidationNode;
        protected XmlNamespaceManager _namespaceManager;
        protected CultureInfo _cultureInfo;

        public void SetupTestData()
        {
            _package = new ExcelPackage();
            _package.Compatibility.IsWorksheets1Based = true;
            _sheet = _package.Workbook.Worksheets.Add("test");
            _cultureInfo = new CultureInfo("en-US");
        }

        public void CleanupTestData()
        {
            if (_package != null) _package.Dispose();
            _package = null;
            _sheet = null;
            _namespaceManager = null;
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            _namespaceManager.AddNamespace("xr", "urn:b");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\">", validationType, address);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected IExcelDataValidationInt CreateSheetWithIntegerValidation(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("NewSheet");
            var validation = sheet.DataValidations.AddIntegerValidation("A1");
            validation.Formula.Value = 1;
            validation.Formula2.Value = 1;

            return validation;
        }

        protected ExcelPackage ReadPackageAsNewPackage(ExcelPackage package)
        {
            MemoryStream xmlStream = new MemoryStream();
            package.SaveAs(xmlStream);

            return new ExcelPackage(xmlStream);
        }

        protected IExcelDataValidationInt ReadIntValidation(ExcelPackage package)
        {
            return (IExcelDataValidationInt)ReadPackageAsNewPackage(package).Workbook.Worksheets[0].DataValidations[0];
        }

        protected T ReadTValidation<T>(ExcelPackage package)
        {
            var validation = ReadPackageAsNewPackage(package).
                Workbook.Worksheets[0].DataValidations[0];

            return (T)((Object)validation);
        }
    }
}
