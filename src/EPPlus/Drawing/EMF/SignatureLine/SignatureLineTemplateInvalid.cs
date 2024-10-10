using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateInvalid : SignatureLineTemplateEmf
    {
        const string invalidPath = "C:\\epplusTest\\Testoutput\\InvalidSignatureTemplate.emf";

        private EMR_EXTTEXTOUTW signedBy;
        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        internal SignatureLineTemplateInvalid(): base(invalidPath)
        {
            var clipRect = (EMR_INTERSECTCLIPRECT)records[128];
            clipRect.Clip.Left = 41;
            clipRect.Clip.Top = 51;
            clipRect.Clip.Right = 242;
            clipRect.Clip.Bottom = 72;
        }

        internal override void SetProperties()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
        }
    }
}
