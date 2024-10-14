using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class SignatureLineEmf : EmfImage
    {
        const string templatePath = @"C:\Users\OssianEdström\Documents\Epplus_Repos\EpplusDigSig\src\EPPlus\resources\SignatureLineTemplate.emf";
        List<EMR_EXTTEXTOUTW> textObjects = new List<EMR_EXTTEXTOUTW>();
        ZipPackagePart part;

        internal string SignerName;
        internal string SignerTitle;

        internal SignatureLineEmf() : base()
        {
            Init();
            Read(templatePath);
        }

        internal SignatureLineEmf(string signerName, string signerTitle) : base()
        {
            SignerName = signerName;
            SignerTitle = signerTitle;
            Init();
        }

        void Init()
        {
            Read(templatePath);
            var aRecord = records;

            var textRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(1);

            foreach (var record in textRecordArr)
            {
                textObjects.Add((EMR_EXTTEXTOUTW)record);
            }
        }

        public override void SaveToStream(MemoryStream ms)
        {
            if (textObjects.Count > 0)
            {
                textObjects[0].Text = SignerName;
            }
            if (textObjects.Count > 1)
            {
                textObjects[1].Text = SignerTitle;
            }
            base.SaveToStream(ms);
        }
    }
}
