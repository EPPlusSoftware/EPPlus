using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateInvalid : SignatureLineTemplateEmf
    {
        private EMR_EXTTEXTOUTW signedBy;
        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        //Template contains records for both valid and invalid files
        //Remove the ones for the Valid template.
        private void RemoveValidRecords()
        {
            for (int i = 62; i <= 69; i++)
            {
                //A removed record automatically makes the next take its place
                records.Remove(records[62]);
            }
        }

        internal SignatureLineTemplateInvalid(SignatureLineEmf sLine) : base(sLine)
        {
            InitalizeClipRect((EMR_INTERSECTCLIPRECT)records[128]);
        }

        internal SignatureLineTemplateInvalid(): base()
        {
            InitalizeClipRect((EMR_INTERSECTCLIPRECT)records[128]);
        }

        internal override void Initalize()
        {
            RemoveValidRecords();
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
        }
    }
}
