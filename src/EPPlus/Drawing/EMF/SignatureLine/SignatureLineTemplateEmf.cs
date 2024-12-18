using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmf : SignatureLineTemplateEmfBase
    {
        const string templateName = "SignatureLineTemplate.emf";
        internal EMR_EXTTEXTOUTW signTextObject;
        internal EMR_EXTTEXTOUTW signedBy;

        internal string SignText
        {
            set
            {
                signTextObject.Text = AdjustText(25, value);
            }
        }

        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        internal SignatureLineTemplateEmf(byte[] emfBytes) : base(templateName, emfBytes)
        {
            signedBy.Text = "";
            signTextObject.Text = "";
            MaxHeight = 47.2f;
            MaxWidth = 205;
        }

        internal SignatureLineTemplateEmf() : base(templateName)
        {
            signedBy.Text = "";
            signTextObject.Text = "";
            MaxHeight = 47.2f;
            MaxWidth = 205;
        }

        internal override List<EMR_EXTTEXTOUTW> Initalize()
        {
            var textRecords = base.Initalize();

            //Index 1 is 'X' and need not be changed.
            signTextObject = textRecords[2];
            suggestedSignerObject = textRecords[3];
            suggestedTitleObject = textRecords[4];
            signedBy = textRecords[5];

            SetImageRecordMax(MaxHeight, MaxWidth);

            return textRecords;
        }

        internal override void SaveTemplateProperties(string[] arr)
        {
            SignedBy = arr[0];
            SignText = arr[1];
        }

        internal override void SaveImage(byte[] imageBytes)
        {
            base.SaveImage(imageBytes);
            imageRecord.ReadBmpAndUpdateImage(imageBytes, false, true);
        }
    }
}
