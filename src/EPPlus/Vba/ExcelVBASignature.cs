/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System.Security.Cryptography.Pkcs;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.VBA.Signatures;
using OfficeOpenXml.Vba.Signatures;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.VBA;

/// <summary>
/// The VBA project's code signature properties
/// </summary>
public class ExcelVbaSignature
{
    internal ExcelVbaSignature(ZipPackagePart vbaPart)
    {
        this._vbaPart = vbaPart;
        this.LegacySignature = new ExcelSignatureVersion(new EPPlusVbaSignatureLegacy(vbaPart), VbaSignatureHashAlgorithm.MD5);
        this.AgileSignature = new ExcelSignatureVersion(new EPPlusVbaSignatureAgile(vbaPart), VbaSignatureHashAlgorithm.SHA1);
        this.V3Signature = new ExcelSignatureVersion(new EPPlusVbaSignatureV3(vbaPart), VbaSignatureHashAlgorithm.SHA1);

        if (this.LegacySignature.Certificate != null)
        {
            this._certificate = this.LegacySignature.Certificate;
        }

        if (this._certificate == null && this.AgileSignature.Certificate != null)
        {
            this._certificate = this.AgileSignature.Certificate;
        }

        if (this._certificate == null && this.V3Signature.Certificate != null)
        {
            this._certificate = this.V3Signature.Certificate;
        }
    }

    internal readonly ZipPackagePart _vbaPart = null;
    private X509Certificate2 _certificate;

    /// <summary>
    /// The certificate to sign the VBA project.
    /// <remarks>
    /// This certificate must have a private key.
    /// There is no validation that the certificate is valid for codesigning, so make sure it's valid to sign Excel files (Excel 2010 is more strict that prior versions).
    /// </remarks>
    /// </summary>
    public X509Certificate2 Certificate
    {
        get { return this._certificate; }
        set
        {
            if (this._certificate == null
                && value != null
                && this.LegacySignature.CreateSignatureOnSave == false
                && this.AgileSignature.CreateSignatureOnSave == false
                && this.V3Signature.CreateSignatureOnSave == false)
            {
                //If we set a new certificate, make sure all signatures are written by default.
                this.LegacySignature.CreateSignatureOnSave = true;
                this.AgileSignature.CreateSignatureOnSave = true;
                this.V3Signature.CreateSignatureOnSave = true;
            }

            this.LegacySignature.Certificate = value;
            this.AgileSignature.Certificate = value;
            this.V3Signature.Certificate = value;
            this._certificate = value;
        }
    }

    /// <summary>
    /// The verifier (legacy format)
    /// </summary>
    public SignedCms Verifier { get; internal set; }

    internal CompoundDocument Signature { get; set; }

    internal ZipPackagePart Part { get; set; }

    internal void Save(ExcelVbaProject proj)
    {
        if (this.Certificate == null)
        {
            return;
        }

        //Legacy signature
        if (this.LegacySignature.CreateSignatureOnSave)
        {
            this.LegacySignature.SignatureHandler.Certificate = this.Certificate;
            this.LegacySignature.CreateSignature(proj);
        }
        else if (this.Part?.Uri != null && this.Part.Package.PartExists(this.Part.Uri))
        {
            this.Part.Package.DeletePart(this.Part.Uri);
        }

        //Agile signature
        ZipPackagePart? p = this.AgileSignature.Part;

        if (this.AgileSignature.CreateSignatureOnSave)
        {
            this.AgileSignature.SignatureHandler.Certificate = this.Certificate;
            this.AgileSignature.CreateSignature(proj);
        }
        else if (p?.Uri != null && p.Package.PartExists(p.Uri))
        {
            p.Package.DeletePart(p.Uri);
        }

        //V3 signature
        p = this.V3Signature.Part;

        if (this.V3Signature.CreateSignatureOnSave)
        {
            this.V3Signature.Certificate = this.Certificate;
            this.V3Signature.CreateSignature(proj);
        }
        else if (p?.Uri != null && p.Package.PartExists(p.Uri))
        {
            p.Package.DeletePart(p.Uri);
        }
    }

    /// <summary>
    /// Settings for the legacy signing.
    /// </summary>
    public ExcelSignatureVersion LegacySignature { get; set; }

    /// <summary>
    /// Settings for the agile vba signing. 
    /// The agile signature adds a hash that is calculated for user forms data in the vba project (designer streams). 
    /// </summary>
    public ExcelSignatureVersion AgileSignature { get; set; }

    /// <summary>
    /// Settings for the V3 vba signing.
    /// The V3 signature includes more coverage for data in the dir and project stream in the hash, not covered by the legacy and agile signatures.
    /// </summary>
    public ExcelSignatureVersion V3Signature { get; set; }
}