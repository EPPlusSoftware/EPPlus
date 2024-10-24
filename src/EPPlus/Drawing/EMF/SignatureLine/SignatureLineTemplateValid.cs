using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateValid : SignatureLineTemplateEmf
    {
        internal EMR_EXTTEXTOUTW timeStamp;
        internal EMR_EXTTEXTOUTW signedBy;
        const string filePath = "C:\\epplusTest\\Testoutput\\ValidSignatureTemplate.emf";

        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        internal SignatureLineTemplateValid() : base(filePath)
        {
            var clipRect = (EMR_INTERSECTCLIPRECT)records[121];
            clipRect.Clip.Left = 41;
            clipRect.Clip.Top = 51;
            clipRect.Clip.Right = 242;
            clipRect.Clip.Bottom = 72;
        }

        internal override void SetProperties()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            timeStamp = (EMR_EXTTEXTOUTW)textRecords[0];
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
        }
    }
}
