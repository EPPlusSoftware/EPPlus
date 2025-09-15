/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Font;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelParagraphBullet : XmlHelper, IPictureContainer
    {
        string _path;
        Action _initXml;
        IPictureRelationDocument _prd;
        internal ExcelParagraphBullet(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nsm, XmlNode topNode, string path, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode)
        {
            _prd = pictureRelationDocument;
            _initXml = initXml;
            _path = path;
            Color = new ExcelDrawingColorManager(nsm, topNode, path + "/a:buClr", SchemaNodeOrder, initXml);
            Font = new ExcelDrawingFontSpecial(nsm, topNode, path + "/a:buFont", initXml);
            Size = new ExcelDrawingBulletSize(nsm, topNode, path, SchemaNodeOrder, initXml);
            AutoNumberType = GetXmlEnumNull<eBulletAutoNumberType>(path + "/a:buAutoNum/@type");
            if(AutoNumberType.HasValue)
            {
                BulletType = eBulletType.AutoNum;
                StartAt = GetXmlNodeInt(path + "/a:buAutoNum/@startAt", 1);
            }
            else
            {
                if(ExistsNode(path + "/a:buNone"))
                {
                    BulletType = eBulletType.None;
                }
                else
                {
                    string buChar = GetXmlNodeString(path + "/a:buChar/@char");
                    if(string.IsNullOrEmpty(buChar))
                    {
                        RelId = GetXmlNodeString(path + "/a:buBlip/@r:embed");
                        if(string.IsNullOrEmpty(RelId)==false)
                        {
                            BulletType = eBulletType.Blip;
                            var imageBytes = PictureStore.GetPicture(RelId, this, out string contentType, out ePictureType pt);
                            BulletImage = new ExcelImage(imageBytes, pt);
                            //TODO: Add 
                        }
                        else
                        {
                            BulletType = eBulletType.None;
                        }
                    }
                    else
                    {
                        BulletType = eBulletType.Character;
                        BulletCharacter = buChar[0];
                    }
                }
            }
        }
        /// <summary>
        /// The type of bullet.
        /// </summary>
        public eBulletType BulletType
        {
            get;
            private set;
        }
        /// <summary>
        /// The color of the bullet character.
        /// </summary>
        public ExcelDrawingColorManager Color { get;  }
        /// <summary>
        /// The font used for the bullet character.
        /// </summary>
        public ExcelDrawingFontSpecial Font { get; }
        /// <summary>
        /// The size used for the bullet.
        /// </summary>
        public ExcelDrawingBulletSize Size { get; }
        /// <summary>
        /// The bullet character, if bullet type is set to Character.
        /// </summary>
        public char? BulletCharacter { get; private set; } 
        /// <summary>
        /// The start value if the bullet type is set to AutoNumber. Default is 1.
        /// </summary>
        public int StartAt { get; private set; }
        /// <summary>
        /// The type of auto numbering, if bullet type is set to AutoNumber.
        /// </summary>
        public eBulletAutoNumberType? AutoNumberType 
        {
            get;
            private set;
        }
        ExcelImage _image=null;
        /// <summary>
        /// The BulletImage
        /// </summary>
        public ExcelImage BulletImage
        {
            get
            {
                if (BulletType != eBulletType.Blip) return null;
                return _image;
            }
            private set
            {
                _image = value;
            }
        }
        private string RelId { get; set; }
        IPictureRelationDocument IPictureContainer.RelationDocument => _prd;

        string IPictureContainer.ImageHash { get; set; }
        Uri IPictureContainer.UriPic { get; set; }
        ZipPackageRelationship IPictureContainer.RelPic { get; set; }
        /// <summary>
        /// Sets the Bullet Type to None
        /// </summary>
        public void SetNone()
        {
            BulletType = eBulletType.None;
            AutoNumberType = null;
            BulletCharacter = null;
            CreateNode(_path + "/a:buNone");
            DeleteNode(_path + "/a:buChar");
            DeleteNode(_path + "/a:buAutoNum");
            DeleteNode(_path + "/a:buBlip");
        }
        public void SetAutoNumber(eBulletAutoNumberType type, int startAt = 1)
        {            
            AutoNumberType = type;
            StartAt = startAt;
            SetXmlNodeString(_path + "/a:buAutoNum/@type", type.ToEnumString());
            if (startAt > 1)
            {
                SetXmlNodeInt(_path + "/a:buAutoNum/@startAt", startAt);
            }
            DeleteNode(_path + "/a:buNone");
            DeleteNode(_path + "/a:buChar");
            DeleteNode(_path + "/a:buBlip");
            BulletType = eBulletType.AutoNum;
        }
        public void SetCharacter(char bulletCharacter)
        {
            DeleteNode(_path + "/a:buNone");
            DeleteNode(_path + "/a:buAutoNum");
            DeleteNode(_path + "/a:buBlip");

            SetXmlNodeString(_path + "/a:buChar", bulletCharacter.ToString());
            BulletType = eBulletType.Character;
        }
        public void SetPicture(Stream image)
        {
            CreateNode(_path + "/a:buNone");
            DeleteNode(_path + "/a:buChar");
            DeleteNode(_path + "/a:buAutoNum");
            DeleteNode(_path + "/a:buBlip");
            if (BulletType==eBulletType.Blip)
            {
                PictureStore.RemoveImage(this);
            }
            BulletImage = new ExcelImage(image);
            ((IPictureContainer)this).SetNewImage();
            BulletType = eBulletType.Blip;
        }

        void IPictureContainer.RemoveImage()
        {
            PictureStore.RemoveImage(this);
        }

        void IPictureContainer.SetNewImage()
        {
            SetXmlNodeString(_path + "/a:buBlip/@r:embed", ((IPictureContainer)this).RelPic.Id);
        }
    }
}