using System;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;

/// <summary>
/// Helpers for building .NET public-key objects from JWK (JSON Web Key) JSON,
/// as supplied by a counterpart for JOSE encrypt/decrypt operations.
/// </summary>
internal static class JwkHelper
{
    /// <summary>
    /// Builds an RSA public key (for RSA-OAEP-256 CEK wrapping) from an RSA JWK JSON string
    /// (expects "n" and "e" members, base64url-encoded per RFC 7518).
    /// </summary>
    public static RSA RsaPublicKeyFromJwk(string jwkJson)
    {
        JObject jwk = JObject.Parse(jwkJson);
        string n = (string)jwk["n"];
        string e = (string)jwk["e"];

        if (string.IsNullOrEmpty(n) || string.IsNullOrEmpty(e))
        {
            throw new InvalidOperationException("RSA JWK is missing required 'n' or 'e' member.");
        }

        var parameters = new RSAParameters
        {
            Modulus = Base64UrlEncoding.Decode(n),
            Exponent = Base64UrlEncoding.Decode(e)
        };

        var rsa = new RSACng();
        rsa.ImportParameters(parameters);
        return rsa;
    }

    /// <summary>
    /// Builds an EC (P-256) public key (for ES256 signature verification) from an EC JWK
    /// JSON string (expects "crv", "x" and "y" members, base64url-encoded per RFC 7518).
    /// </summary>
    /// <remarks>
    /// .NET Framework 4.6.2 does not expose the ECParameters/ECPoint/ECCurve API surface
    /// (added in .NET Framework 4.7). Instead, the public key is imported as a raw CNG
    /// ECC public key blob (BCRYPT_ECCKEY_BLOB) via <see cref="CngKey"/>, wrapped in an
    /// <see cref="ECDsaCng"/> instance (which derives from <see cref="ECDsa"/>).
    /// </remarks>
    public static ECDsa EcPublicKeyFromJwk(string jwkJson)
    {
        JObject jwk = JObject.Parse(jwkJson);
        string crv = (string)jwk["crv"];
        string x = (string)jwk["x"];
        string y = (string)jwk["y"];

        if (string.IsNullOrEmpty(x) || string.IsNullOrEmpty(y))
        {
            throw new InvalidOperationException("EC JWK is missing required 'x' or 'y' member.");
        }

        if (!string.IsNullOrEmpty(crv) && crv != "P-256")
        {
            throw new InvalidOperationException($"Unsupported EC curve '{crv}'. Only P-256 (ES256) is supported.");
        }

        const int keySizeBytes = 32; // P-256
        const uint bcryptEcdsaPublicP256Magic = 0x31534345; // "ECS1"

        byte[] xBytes = PadLeft(Base64UrlEncoding.Decode(x), keySizeBytes);
        byte[] yBytes = PadLeft(Base64UrlEncoding.Decode(y), keySizeBytes);

        byte[] blob = new byte[8 + (keySizeBytes * 2)];
        BitConverter.GetBytes(bcryptEcdsaPublicP256Magic).CopyTo(blob, 0);
        BitConverter.GetBytes(keySizeBytes).CopyTo(blob, 4);
        Array.Copy(xBytes, 0, blob, 8, keySizeBytes);
        Array.Copy(yBytes, 0, blob, 8 + keySizeBytes, keySizeBytes);

        using (CngKey cngKey = CngKey.Import(blob, CngKeyBlobFormat.EccPublicBlob))
        {
            return new ECDsaCng(cngKey);
        }
    }

    private static byte[] PadLeft(byte[] value, int length)
    {
        if (value.Length == length)
        {
            return value;
        }

        if (value.Length > length)
        {
            throw new InvalidOperationException("EC coordinate value is longer than expected for the P-256 curve.");
        }

        byte[] padded = new byte[length];
        Array.Copy(value, 0, padded, length - value.Length, value.Length);
        return padded;
    }
}
