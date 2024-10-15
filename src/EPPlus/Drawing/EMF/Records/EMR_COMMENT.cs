using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_COMMENT : EMR_RECORD
    {
        uint dataSize;
        uint commentIdentifier;
        byte[] PrivateData;

        enum CommentIdentifier
        {
            EMR_COMMENT_EMFSPOOL = 0,
            EMR_COMMENT_EMFPLUS = 0x2B464D45,
            EMR_COMMENT_PUBLIC = 0x43494447
        }

        CommentIdentifier? commentType = null;

        //EMFSpool
        uint EMFSpoolRecordIdentifier;
        byte[] EMFSpoolRecords;

        //EMFPLUS
        byte[] EMFPLUSRECORDS;

        //PUBLIC
        uint PublicCommentIdentifier;

        enum EmrComment : uint
        {
            EMR_COMMENT_WINDOWS_METAFILE = 0x80000001,
            EMR_COMMENT_BEGINGROUP = 0x00000002,
            EMR_COMMENT_ENDGROUP = 0x00000003,
            EMR_COMMENT_MULTIFORMATS = 0x40000004,
            EMR_COMMENT_UNICODE_STRING = 0x00000040,
            EMR_COMMENT_UNICODE_END = 0x00000080
        }

        EmrComment emrComment;
        //BeginComment
        RectLObject rect;
        uint nDescription;
        string Description;

        internal EMR_COMMENT(BinaryReader br) : base(br, (uint)RECORD_TYPES.EMR_COMMENT)
        {
            dataSize = br.ReadUInt32();
            commentIdentifier = br.ReadUInt32();

            switch (commentIdentifier)
            {
                case 0:
                case 0x2B464D45:
                case 0x43494447:
                    commentType = (CommentIdentifier)commentIdentifier;
                    break;
                default:
                    commentType = null;
                    break;
            }

            if (commentType == null)
            {
                br.BaseStream.Position = br.BaseStream.Position - 4;
                PrivateData = new byte[dataSize];
                br.Read(PrivateData, 0, (int)dataSize);
            }
            else if (commentType == CommentIdentifier.EMR_COMMENT_EMFSPOOL)
            {
                EMFSpoolRecordIdentifier = br.ReadUInt32();

                //If value equals TONE in ASCII
                if (EMFSpoolRecordIdentifier == 0x544F4E4)
                {
                    //Handle EMFSPOOL font definition
                }

                EMFSpoolRecords = new byte[dataSize];
                br.Read(PrivateData, 0, (int)dataSize);
            }
            else if (commentType == CommentIdentifier.EMR_COMMENT_EMFPLUS)
            {
                EMFPLUSRECORDS = new byte[dataSize];
                br.Read(PrivateData, 0, (int)dataSize);
            }
            else if (commentType == CommentIdentifier.EMR_COMMENT_PUBLIC)
            {
                PublicCommentIdentifier = br.ReadUInt32();
                emrComment = (EmrComment)PublicCommentIdentifier;

                if (emrComment == EmrComment.EMR_COMMENT_BEGINGROUP)
                {
                    rect = new RectLObject(br);
                    nDescription = br.ReadUInt32();
                    Description = BinaryHelper.GetString(br, (nDescription * 2), Encoding.Unicode);

                }
            }
        }
    }
}
