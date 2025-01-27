using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmfStamp : SignatureLineTemplateEmfBase
    {
        const string templateName = "SignatureLineStampTemplate.emf";

        const int minWidth = 10;
        const int maxWidthFile = 127;

        internal override string SuggestedSigner
        {
            set
            {
                base.SuggestedSigner = value;
                suggestedSignerObject.AdjustReferenceToCenterText(maxWidthFile, minWidth);
            }
        }

        internal override string TimeStamp
        {
            set
            {
                base.TimeStamp = value;
                timeStamp.AdjustReferenceToCenterText(maxWidthFile, minWidth);
            }
        }

        internal override string SuggestedTitle
        {
            set
            {
                base.SuggestedTitle = value;
                suggestedTitleObject.AdjustReferenceToCenterText(maxWidthFile, minWidth);
            }
            get
            {
                return suggestedTitleObject.Text;
            }
        }


        internal SignatureLineTemplateEmfStamp(byte[] emfBytes) : base(templateName, emfBytes)
        {
            MaxHeight = 75;
            MaxWidth = 111;
        }

        internal SignatureLineTemplateEmfStamp() : base(templateName)
        {
            MaxHeight = 75;
            MaxWidth = 111;
        }

        internal override List<EMR_EXTTEXTOUTW> Initalize()
        {
            var textRecords = base.Initalize();

            suggestedTitleObject = textRecords[1];
            suggestedSignerObject = textRecords[2];

            SetImageRecordMax(MaxHeight, MaxWidth);

            return textRecords;
        }

        internal override void SaveTemplateProperties(string[] arr)
        {
            SuggestedSigner = arr[0];
        }

        internal override void SaveImage(byte[] imageBytes)
        {
            base.SaveImage(imageBytes);
            imageRecord.ReadBmpAndUpdateImage(imageBytes, true, false);
        }
    }
}
