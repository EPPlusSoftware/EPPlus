/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/11/2024         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils.RemoteCalls;
using OfficeOpenXml.CellPictures;
using System.IO;
using System.Net;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.LookupAndReference,
       EPPlusVersion = "8",
       Description = "Inserts an image into cell via a https call")]
    internal class ImageFunction : ExcelFunctionAsync
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;

        public override CompileResult Complete(RemoteTask task)
        {
            var ctx = task.ParsingContext;
            // execute the rest of the function here
            return null;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var url = ArgToString(arguments, 0);
            if(string.IsNullOrEmpty(url) || url.Length < 8 || !IsValidHttpsUrl(url))
            {
                return CreateResult(eErrorType.Value);
            }
            var altText = default(string);
            var sizing = default(int?);
            var height = default(double?);
            var width = default(double?);
            if(arguments.Count > 1)
            {
                altText = ArgToString(arguments, 1);
            }
            if (arguments.Count > 2)
            {
                sizing = ArgToInt(arguments, 2, out ExcelErrorValue e1, 0);
                if (e1 != null) return CreateResult(e1, DataType.ExcelError);
                if(sizing < 0 || sizing > 3)
                {
                    return CreateResult(eErrorType.Value);
                }
            }
            if(arguments.Count > 3)
            {
                height = ArgToDecimal(arguments, 3, out ExcelErrorValue e2);
                if (e2 != null) return CreateResult(e2, DataType.ExcelError);
            }
            if(arguments.Count > 4)
            {
                width = ArgToDecimal(arguments, 4, out ExcelErrorValue e3);
                if(e3 != null) return CreateResult(e3, DataType.ExcelError);
            }
            var httpTask = new HttpRemoteTask(url, this, context);
            //RemoteCallManager.QueueTask(httpTask);
            // return #BUSY error

            var cellPictureManager = new CellPicturesManager(context.CurrentWorksheet);
            var httpsService = context.CurrentWorksheet._package.Settings.ImageFunctionService;
            if(httpsService == null)
            {
                return CreateResult(eErrorType.Value);
            }
            var imageBytes = httpsService.Download(url);
            cellPictureManager.SetWebPicture(context.CurrentCell.Row, context.CurrentCell.Column, new Uri(url), imageBytes, null);
            var cellPic = context.CurrentWorksheet.GetValue(context.CurrentCell.Row, context.CurrentCell.Column);
            return CreateResult(cellPic, DataType.WebImage);
        }

        private bool IsValidHttpsUrl(string url)
        {  
            try
            {
                var uri = new Uri(url, UriKind.Absolute);
                return uri.Scheme == Uri.UriSchemeHttps;
            }
            catch
            {
                return false;
            }
        }
    }
}
