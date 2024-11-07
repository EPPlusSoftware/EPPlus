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

        internal string SuggestedTitle
        {
            set
            {
                suggestedTitleObject.Text = AdjustText(39, value);
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

        ////Template contains records for both valid and invalid files
        ////Remove the ones for the Invalid template.
        //internal void RemoveInvalidRecords()
        //{
        //    if(IsStamp)
        //    {
        //        //Stamp records have a different structure
        //        for (int i = 51; i <= 68; i++)
        //        {
        //            //A removed record automatically makes the next take its place
        //            records.Remove(records[51]);
        //        }
        //    }
        //    else
        //    {
        //        //We remove records "backwards" so that indicies for the next operation do not change.

        //        for (int i = 69; i <= 75; i++)
        //        {
        //            //A removed record automatically makes the next take its place
        //            records.Remove(records[69]);
        //        }

        //        records.Remove(records[62]);

        //        for (int i = 54; i <= 60; i++)
        //        {
        //            //A removed record automatically makes the next take its place
        //            records.Remove(records[51]);
        //        }
        //    }
        //}

        ////Template contains records for both valid and invalid files
        ////Remove the ones for the Valid template.
        ///
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

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);

            timeStamp = (EMR_EXTTEXTOUTW)textRecords[0];
            //1 is 'Invalid Signature' 2 is 'X' neither need be changed.

            if(IsStamp)
            {
                suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[1];
                suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[2];
                records.RemoveAt(62);
                records.RemoveAt(51);
            }
            else
            {
                signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
                suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
                suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
                signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
            }
            //var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(1).ToArray();
            //signTextObject = (EMR_EXTTEXTOUTW)textRecords[0];
            //suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[1];
            //suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[2];
        }
    }
}
