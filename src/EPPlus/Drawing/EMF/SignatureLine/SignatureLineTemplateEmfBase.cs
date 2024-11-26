using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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

        protected const string localZipPath = "OfficeOpenXml.resources.SignatureLineTemplates.zip";

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
                suggestedTitleObject.Text = value;
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
                return inputString.Substring(0, length-1) + "...";
            }
            return inputString;
        }

        //Load image record from original Emf into template
        internal SignatureLineTemplateEmfBase(string templateName, byte[] originalBytes) : this(templateName)
        {
            var tmp = new EmfImage();
            tmp.Read(originalBytes);
            var tmpImageRecord = (EMR_STRETCHDIBITS)tmp.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            imageRecord.ExtractedBmp = tmpImageRecord.ExtractedBmp;
        }
        internal byte[] GetBitmapBytes()
        {
            return imageRecord.ExtractedBmp.GetBitMapBytes();
        }

        internal SignatureLineTemplateEmfBase(EmfImage emf)
        {
            Read(emf.GetBytes());
            Initalize();
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
            invalidRecords.LoadTemplateFromResource("InvalidSignatureRecords.bin", "OfficeOpenXml.resources.SignatureLineTemplates.zip");
            records.InsertRange(63, invalidRecords.records);
        }

        internal virtual List<EMR_EXTTEXTOUTW> Initalize()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();
            timeStamp = textRecords[0];
            imageRecord = (EMR_STRETCHDIBITS)records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
            return textRecords;
        }

        internal void SetImageRecordMax(float MaxHeight, float MaxWidth)
        {
            imageRecord.MaxHeight = MaxHeight;
            imageRecord.MaxWidth = MaxWidth;
        }

        internal virtual void SaveTemplateProperties(string[] arr)
        {
        }

        internal virtual void SaveImage(byte[] imageBytes)
        {
        }
    }
}
