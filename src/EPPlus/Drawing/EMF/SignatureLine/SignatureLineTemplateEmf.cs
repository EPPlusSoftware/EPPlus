using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmf : EmfImage
    {
        internal EMR_EXTTEXTOUTW timeStamp;
        internal EMR_EXTTEXTOUTW signTextObject;
        internal EMR_EXTTEXTOUTW suggestedSignerObject;
        internal EMR_EXTTEXTOUTW suggestedTitleObject;
        internal EMR_EXTTEXTOUTW signedBy;

        const int minWidth = 10;
        const int maxWidth = 127;

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
                if(IsStamp)
                {
                    suggestedSignerObject.AdjustReferenceToCenterText(maxWidth, minWidth);
                }
            }
        }


        internal string TimeStamp
        {
            set
            {
                timeStamp.Text = AdjustText(39, value);
                if (IsStamp)
                {
                    timeStamp.AdjustReferenceToCenterText(maxWidth, minWidth);
                }
            }
        }

        internal string SuggestedTitle
        {
            set
            {
                suggestedTitleObject.Text = value;
                if(IsStamp)
                {
                    suggestedTitleObject.AdjustReferenceToCenterText(maxWidth, minWidth);
                }
            }
            get
            {
                return suggestedTitleObject.Text;
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

        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        internal SignatureLineTemplateEmf(EmfImage emf)
        {
            Read(emf.GetBytes());
            Initalize();
        }

        internal bool IsStamp = false;

        internal SignatureLineTemplateEmf(bool isStamp = false)
        {
            records.Clear();

            var path = isStamp ? "SignatureLineStampTemplate.emf" : "SignatureLineTemplate.emf";
            LoadTemplateFromResource(path, "OfficeOpenXml.resources.SignatureLineTemplates.zip");
            IsStamp = isStamp;

            Initalize();
            timeStamp.Text = "";
            SuggestedTitle = "Developer";
            Save("C:\\epplusTest\\Testoutput\\LoadedFromZip.emf");
        }

        internal void RemoveValidRecords()
        {
            for (int i = 63; i <= 79; i++)
            {
                //A removed record automatically makes the next take its place
                records.Remove(records[63]);
            }
        }

        internal void InsertInvalidRecords()
        {
            EmfImage invalidRecords = new EmfImage();
            invalidRecords.LoadTemplateFromResource("InvalidSignatureRecords.bin", "OfficeOpenXml.resources.SignatureLineTemplates.zip");
            records.InsertRange(63, invalidRecords.records);
        }

        internal new SignatureLineTemplateEmf Clone()
        {
            var copy = new SignatureLineTemplateEmf(this);
            return copy;
        }


        internal virtual void Initalize()
        {
            var aRecord = records;

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Cast<EMR_EXTTEXTOUTW>().ToList();

            timeStamp = textRecords[0];
            //1 is 'Invalid Signature' 2 is 'X' neither need be changed.

            if(IsStamp)
            {
                suggestedTitleObject = textRecords[1];
                suggestedSignerObject = textRecords[2];
            }
            else
            {
                signTextObject = textRecords[2];
                suggestedSignerObject = textRecords[3];
                suggestedTitleObject = textRecords[4];
                signedBy = textRecords[5];
            }
        }
    }
}
