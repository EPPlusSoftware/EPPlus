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
using System;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Xml;
using System.Globalization;

namespace EPPlusTest.DataValidation
{
    public abstract class ValidationTestBase
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
            _package = null;
            _sheet = null;
            _namespaceManager = null;
        }

        protected string GetTestOutputPath(string fileName)
        {
            return Path.Combine(
#if (Core)
            Path.GetTempPath()      //In Net.Core Output to TempPath 
#else
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
#endif
                , fileName);
        }

        protected void SaveTestOutput(string fileName)
        {
            var path = GetTestOutputPath(fileName);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            _package.SaveAs(new FileInfo(path));
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\">", validationType, address);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected void LoadXmlTestData(string address, string validationType, string operatorName, string formula1Value)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\" operator=\"{2}\">", validationType, address, operatorName);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value, bool showErrorMsg, bool showInputMsg)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\" ", validationType, address);
            sb.AppendFormat(" showErrorMessage=\"{0}\" showInputMessage=\"{1}\">", showErrorMsg ? 1 : 0, showInputMsg ? 1 : 0);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value, string prompt, string promptTitle, string error, string errorTitle)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\"", validationType, address);
            sb.AppendFormat(" prompt=\"{0}\" promptTitle=\"{1}\"", prompt, promptTitle);
            sb.AppendFormat(" error=\"{0}\" errorTitle=\"{1}\">", error, errorTitle);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

    }
}
