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
        List<EMR_EXTTEXTOUTW> textObjects = new List<EMR_EXTTEXTOUTW>();
        ZipPackagePart part;

        internal string SignerName;
        internal string SignerTitle;
        bool IsStamp = false;

        internal SignatureLineEmf() : base()
        {
            InitTemplate();
            Init();
        }

        internal SignatureLineEmf(string signerName, string signerTitle) : base()
        {
            SignerName = signerName;
            SignerTitle = signerTitle;
            InitTemplate();
            Init();
        }

        void Init()
        {
            var aRecord = records;

            var textRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            //textObjects.Add((EMR_EXTTEXTOUTW)textRecordArr[3]);
            textObjects.Add((EMR_EXTTEXTOUTW)textRecordArr[4]);
            textObjects.Add((EMR_EXTTEXTOUTW)textRecordArr[5]);
            //textObjects.Add((EMR_EXTTEXTOUTW)textRecordArr[6]);


            //// var textRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(1);

            //foreach (var record in textRecordArr)
            //{
            //    textObjects.Add((EMR_EXTTEXTOUTW)record);
            //}
        }

        internal void InitTemplate(bool isStamp = false)
        {
            //var currDir = Directory.GetCurrentDirectory();
            //var path = $@"{currDir}\resources\";

            //if (isStamp)
            //{
            //    path += "SignatureLineStampTemplate.emf";
            //}
            //else
            //{
            //    path += "SignatureLineTemplate.emf";
            //}

            records.Clear();

            var path = isStamp ? "SignatureLineStampTemplate.emf" : "SignatureLineTemplate.emf";
            LoadTemplateFromResource(path, "OfficeOpenXml.resources.SignatureLineTemplates.zip");

            //If type has changed, re-init
            if (isStamp != IsStamp)
            {
                Init();
                IsStamp = isStamp;
            }
        }

        //Remove unnecesary records
        private void RemoveRecords(EmfImage image)
        {
            //We remove records "backwards" to avoid confusion around altered indicies
            //e.g. if we removed records[51] first we would then need to remove records[61] instead of [62]
            //Desipte its original index being 62

            image.records.Remove(image.records[181]);
            image.records.Remove(image.records[138]);

            for (int i = 69; i <= 75; i++)
            {
                //A removed record automatically makes the next take its place
                image.records.Remove(image.records[69]);
            }

            for (int i = 62; i <= 69; i++)
            {
                //A removed record automatically makes the next take its place
                image.records.Remove(image.records[62]);
            }

            for (int i = 54; i <= 60; i++)
            {
                //A removed record automatically makes the next take its place
                image.records.Remove(image.records[51]);
            }
        }

        public override void SaveToStream(MemoryStream ms)
        {
            textObjects[0].Text = SignerName;
            textObjects[1].Text = SignerTitle;

            //By using a clone here we don't need to re-create records for valid/invalid
            //If this is part of a signature
            var saveObj = Clone();
            RemoveRecords(saveObj);

            saveObj.SaveToStream(ms);
            saveObj.Save(@"C:\epplusTest\Testoutput\image1Generated.emf");
        }
    }
}
