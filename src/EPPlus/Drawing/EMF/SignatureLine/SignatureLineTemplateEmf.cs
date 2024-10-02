using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmf : EMF
    {
        internal EMR_EXTTEXTOUTW signTextObject;
        internal EMR_EXTTEXTOUTW suggestedSignerObject;
        internal EMR_EXTTEXTOUTW suggestedTitleObject;

        internal string SignText
        {
            set
            {
                signTextObject.Text = AdjustText(25, value);
            }
        }

        internal string SuggestedSigner
        {
            set
            {
                suggestedSignerObject.Text = AdjustText(39, value);
            }
        }

        internal string SuggestedTitle
        {
            set
            {
                suggestedTitleObject.Text = AdjustText(39, value);
            }
        }

        string AdjustText(int length, string inputString)
        {
            if (inputString.Length > length)
            {
                return inputString.Substring(0, length-1) + "...";
            }
            return inputString;
        }

        internal SignatureLineTemplateEmf(string templatePath)
        {
            Read(templatePath);
            SetProperties();
        }

        internal virtual void SetProperties()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(1).ToArray();
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[0];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[1];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[2];
        }
    }
}
