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
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.RichData.RichValues.WebImages;
using OfficeOpenXml.Utils.RemoteCalls;
using System;
using System.Collections.Generic;

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
            var sizing =WebImageSizing.FitToCellMaintainRatio;
            var height = default(double?);
            var width = default(double?);
            if(arguments.Count > 1)
            {
                altText = ArgToString(arguments, 1);
            }
            if (arguments.Count > 2)
            {
                var sizingInt = ArgToInt(arguments, 2, out ExcelErrorValue e1, 0);
                if (e1 != null) return CreateResult(e1, DataType.ExcelError);
                if(sizingInt < 0 || sizingInt > 3)
                {
                    return CreateResult(eErrorType.Value);
                }
                sizing = (WebImageSizing)sizingInt;
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

            if(sizing != WebImageSizing.CustomizeByHeightAndWidth && (height.HasValue || width.HasValue))
            {
                return CreateResult(eErrorType.Value);
            }
            else if(sizing == WebImageSizing.CustomizeByHeightAndWidth && !height.HasValue && !width.HasValue)
            {
                return CreateResult(eErrorType.Value);
            }

            // TODO: cache image url:s and internal image Uri (in the picture store) and make http-call only if the image
            // is already present in the workbook.

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
            cellPictureManager.SetWebPicture(context.CurrentCell.Row, context.CurrentCell.Column, new Uri(url), imageBytes, altText, CalcOrigins.Formula, sizing, height, width);
            var cellPic = context.CurrentWorksheet.GetValue(context.CurrentCell.Row, context.CurrentCell.Column);
            //return CreateResult(cellPic, DataType.WebImage);
            return new WebImageCompileResult(cellPic);
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
