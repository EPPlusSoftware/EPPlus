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

        internal string SignedBy
        {
            set
            {
                signedBy.Text = $"Signed by:{value}";
            }
        }

        internal SignatureLineTemplateValid(SignatureLineEmf sLine) : base(sLine)
        {
            InitalizeClipRect((EMR_INTERSECTCLIPRECT)records[121]);
        }

        internal SignatureLineTemplateValid() : base()
        {
            InitalizeClipRect((EMR_INTERSECTCLIPRECT)records[121]);
        }

        //Template contains records for both valid and invalid files
        //Remove the ones for the Invalid template.
        private void RemoveInvalidRecords()
        {
            //We remove records "backwards" so that indicies for the next operation do not change.

            for (int i = 69; i <= 75; i++)
            {
                //A removed record automatically makes the next take its place
                records.Remove(records[69]);
            }

            records.Remove(records[62]);

            for (int i = 54; i <= 60; i++)
            {
                //A removed record automatically makes the next take its place
                records.Remove(records[51]);
            }
        }

        internal override void Initalize()
        {
            RemoveInvalidRecords();
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            timeStamp = (EMR_EXTTEXTOUTW)textRecords[0];
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
        }
    }
}
