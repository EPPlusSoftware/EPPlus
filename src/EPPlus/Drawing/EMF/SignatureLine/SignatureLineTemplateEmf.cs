using System.Linq;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineTemplateEmf : EmfImage
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
        protected void InitalizeClipRect(EMR_INTERSECTCLIPRECT clipRect)
        {
            clipRect.Clip.Left = 41;
            clipRect.Clip.Top = 51;
            clipRect.Clip.Right = 242;
            clipRect.Clip.Bottom = 72;
        }

        internal SignatureLineTemplateEmf(SignatureLineEmf emf)
        {
            Read(emf.GetBytes());
            Initalize();
            SuggestedSigner = emf.SignerName;
            SuggestedTitle = emf.SignerTitle;
        }

        internal SignatureLineTemplateEmf()
        {
            LoadTemplateFromResource("SignatureLineTemplate.emf", "OfficeOpenXml.resources.SignatureLineTemplates.zip");
            Initalize();
        }

        internal virtual void Initalize()
        {
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(1).ToArray();
            signTextObject = (EMR_EXTTEXTOUTW)textRecords[0];
            suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[1];
            suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[2];
        }
    }
}
