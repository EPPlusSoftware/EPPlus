using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmfBase : EmfImage
    {
        internal EMR_EXTTEXTOUTW timeStamp;
        internal EMR_EXTTEXTOUTW suggestedSignerObject;
        internal EMR_EXTTEXTOUTW suggestedTitleObject;
        internal EMR_STRETCHDIBITS imageRecord;
        internal EmfImage EmfSignatureImage = null;

        protected const string localZipPath = "OfficeOpenXml.resources.SignatureLineTemplates.zip";

        protected double MaxHeight = 47.2f;
        protected double MaxWidth = 205;

        internal virtual string SuggestedSigner
        {
            set
            {
                suggestedSignerObject.Text = AdjustText(39, value);
            }
        }

        internal virtual string TimeStamp
        {
            set
            {
                timeStamp.Text = AdjustText(39, value);
            }
        }

        internal virtual string SuggestedTitle
        {
            set
            {
                suggestedTitleObject.Text = AdjustText(39, value);
            }
            get
            {
                return suggestedTitleObject.Text;
            }
        }

        protected string AdjustText(int length, string inputString)
        {
            if (inputString.Length > length)
            {
                return inputString.Substring(0, length - 1) + "...";
            }
            return inputString;
        }

        //Load image record from original Emf into template
        internal SignatureLineTemplateEmfBase(string templateName, byte[] originalBytes) : this(templateName)
        {
            EmfSignatureImage = new EmfImage();
            EmfSignatureImage.Read(originalBytes);
            var tmpImageRecord = (EMR_STRETCHDIBITS)EmfSignatureImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            imageRecord.ExtractedBmp = tmpImageRecord.ExtractedBmp;
        }
        internal byte[] GetBitmapBytes()
        {
            return imageRecord.ExtractedBmp.GetBitMapBytes();
        }

        internal SignatureLineTemplateEmfBase(string templateName)
        {
            records.Clear();
            LoadTemplateFromResource(templateName, localZipPath);
            Initalize();
            timeStamp.Text = "";
        }

        internal void InsertInvalidRecords()
        {
            EmfImage invalidRecords = new EmfImage();
            invalidRecords.LoadTemplateFromResource("InvalidSignatureRecords.bin", localZipPath);
            records.InsertRange(63, invalidRecords.records);
        }

        internal virtual List<EMR_EXTTEXTOUTW> Initalize()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();
            timeStamp = textRecords[0];
            imageRecord = (EMR_STRETCHDIBITS)records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            return textRecords;
        }

        internal void SetImageRecordMax(double MaxHeight, double MaxWidth)
        {
            imageRecord.MaxHeight = MaxHeight;
            imageRecord.MaxWidth = MaxWidth;
        }

        internal virtual void SaveTemplateProperties(string[] arr)
        {
        }

        internal virtual void SaveImage(byte[] imageBytes)
        {
            EmfSignatureImage = new EmfImage();
            EmfSignatureImage.LoadTemplateFromResource("SignatureImageTemplate.emf", localZipPath);

            var imgRecord = (EMR_STRETCHDIBITS)EmfSignatureImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            BitmapHandler handler = new BitmapHandler(imageBytes);
            var infoHeader = handler.informationHeader;

            ImageUtil.ResizeImageWithMaxSize(
                MaxWidth, MaxHeight,
                infoHeader.pixelWidth, infoHeader.pixelHeight,
                out int newWidth, out int newHeight
                );

            imgRecord.MaxWidth = newWidth;
            imgRecord.MaxHeight = newHeight;

            imgRecord.ReadBmpAndUpdateImage(imageBytes);

            var header = (EMR_HEADER)EmfSignatureImage.records[0];

            //Update Bounds
            header.Bounds.Right = newWidth;
            header.Bounds.Bottom = newHeight;

            //Update frame (*100 because unit is in 0.01 mm)
            header.Frame.Right = Convert.ToInt32(header.MilimetersPerPixelX * newWidth * 100);
            header.Frame.Bottom = Convert.ToInt32(header.MilimetersPerPixelY * newHeight * 100);
        }
    }
}
