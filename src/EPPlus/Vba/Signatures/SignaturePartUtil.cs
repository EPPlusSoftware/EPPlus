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
        internal static ZipPackagePart GetPart(ExcelVbaProject proj, EPPlusVbaSignature signature, string schemaRelation)
        {
            var rel = (ZipPackageRelationship)proj.Part.GetRelationshipsByType(schemaRelation).FirstOrDefault();
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
                    // TODO: replace hardcoded signature
                    uri = new Uri("/xl/vbaProjectSignature.bin", UriKind.Relative);
                    part = proj._pck.CreatePart(uri, ContentTypes.contentTypeVBASignature);
                }
            }
            if (rel == null)
            {
                proj.Part.CreateRelationship(UriHelper.ResolvePartUri(proj.Uri, uri), Packaging.TargetMode.Internal, schemaRelation);
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
    }
}
