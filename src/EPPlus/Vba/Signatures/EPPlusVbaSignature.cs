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

using System;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeOpenXml.VBA.Signatures;

internal abstract class EPPlusVbaSignature
{
    public EPPlusVbaSignature(ZipPackagePart vbaPart, ExcelVbaSignatureType signatureType)
    {
        this._vbaPart = vbaPart;
        this._signatureType = signatureType;
        this.Context=new EPPlusSignatureContext(signatureType);
    }

    private readonly ZipPackagePart _vbaPart;
    private readonly ExcelVbaSignatureType _signatureType;
    internal ZipPackagePart Part
    {
        get;
        set;
    }

    internal string SchemaRelation
    {
        get
        {
            switch(this._signatureType)
            {
                case ExcelVbaSignatureType.Legacy:
                    return VbaSchemaRelations.Legacy;
                case ExcelVbaSignatureType.Agile:
                    return VbaSchemaRelations.Agile;
                case ExcelVbaSignatureType.V3:
                    return VbaSchemaRelations.V3;
                default:
                    return VbaSchemaRelations.Legacy;
            }
        }
    }
    internal string ContentType
    {
        get
        {
            switch (this._signatureType)
            {
                case ExcelVbaSignatureType.Legacy:
                    return ContentTypes.contentTypeVBASignature;
                case ExcelVbaSignatureType.Agile:
                    return ContentTypes.contentTypeVBASignatureAgile;
                default:
                    return ContentTypes.contentTypeVBASignatureV3;
            }
        }
    }

    public X509Certificate2 Certificate { get; set; }
    public SignedCms Verifier { get; internal set; }

    public EPPlusSignatureContext Context { get; set; }

    internal bool ReadSignature()
    {

        if (this._vbaPart == null)
        {
            return true; //If no vba part exists, create the signature by default.
        }

        ZipPackageRelationship? rel = this._vbaPart.GetRelationshipsByType(this.SchemaRelation).FirstOrDefault();
        if(rel != null)
        {
            Uri? uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            this.Part = this._vbaPart.Package.GetPart(uri);
            this.Context = new EPPlusSignatureContext(this._signatureType);
            SignatureInfo? signature = SignatureReader.ReadSignature(this.Part, this._signatureType, this.Context);
            this.Certificate = signature.Certificate;
            this.Verifier = signature.Verifier;                
            return true;
        }
        else
        {
            this.Certificate = null;
            this.Verifier = null;
            this.Context = new EPPlusSignatureContext(this._signatureType);
            return false;
        }
    }

    internal void CreateSignature(ExcelVbaProject project)
    {
        byte[] certStore = CertUtil.GetSerializedCertStore(this.Certificate.RawData);
        if (this.Certificate == null)
        {
            SignaturePartUtil.DeleteParts(this.Part);
            return;
        }

        if (this.Certificate.HasPrivateKey == false)    //No signature. Remove any Signature part
        {
            X509Certificate2? storeCert = CertUtil.GetCertificate(this.Certificate.Thumbprint);
            if (storeCert != null)
            {
                this.Certificate = storeCert;
            }
            else
            {
                SignaturePartUtil.DeleteParts(this.Part);
                return;
            }
        }
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter? bw = new BinaryWriter(ms);
        this.Verifier = CertUtil.SignProject(project, this, this.Context);
        byte[]? cert = this.Verifier.Encode();
        byte[]? signatureBytes = CertUtil.CreateBinarySignature(ms, bw, certStore, cert);
        this.Part = SignaturePartUtil.GetPart(project, this);
        this.Part.GetStream(FileMode.Create).Write(signatureBytes, 0, signatureBytes.Length);

    }
}