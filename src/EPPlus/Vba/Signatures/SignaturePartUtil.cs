using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
