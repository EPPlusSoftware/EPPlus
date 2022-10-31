/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Linq;

namespace OfficeOpenXml.VBA.Signatures
{
    internal static class SignaturePartUtil
    {
        internal static ZipPackagePart GetPart(ExcelVbaProject proj, EPPlusVbaSignature signature)
        {
            var rel = (ZipPackageRelationship)proj.Part.GetRelationshipsByType(signature.SchemaRelation).FirstOrDefault();
            var part = signature.Part;
            var uri = default(Uri);
            if (part == null)
            {
                if (rel != null)
                {
                    uri = rel.TargetUri;
                    part = proj._pck.GetPart(rel.TargetUri);
                }
                else
                {
                    uri = GetUriByType(signature.Context.SignatureType, UriKind.Relative);
                    
                    part = proj._pck.CreatePart(uri, signature.ContentType);
                }
            }
            if (rel == null)
            {
                proj.Part.CreateRelationship(UriHelper.ResolvePartUri(proj.Uri, uri), Packaging.TargetMode.Internal, signature.SchemaRelation);
            }
            return part;
        }
        internal static void DeleteParts(params ZipPackagePart[] parts)
        {
            foreach (var part in parts)
            {
                if (part != null)
                {
                    DeletePartAndRelations(part);
                }
            }
        }

        internal static void DeletePartAndRelations(ZipPackagePart part)
        {
            if (part == null) return;
            foreach (var r in part.GetRelationships())
            {
                part.DeleteRelationship(r.Id);
            }
            part.Package.DeletePart(part.Uri);
        }

        private static Uri GetUriByType(ExcelVbaSignatureType signatureType, UriKind uriKind)
        {
            switch (signatureType)
            {
                case ExcelVbaSignatureType.Agile:
                    return new Uri("/xl/vbaProjectSignatureAgile.bin", uriKind);
                case ExcelVbaSignatureType.V3:
                    return new Uri("/xl/vbaProjectSignatureV3.bin", uriKind);
                default:
                    return new Uri("/xl/vbaProjectSignature.bin", uriKind);
            }
        }
    }
}
