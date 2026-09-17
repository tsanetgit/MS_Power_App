using Microsoft.Xrm.Sdk;
using System;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// The plaintext is signed (ES256) with this member's private signing key — held in Azure
/// Key Vault and never exposed locally — and the resulting JWS is then encrypted (A256GCM,
/// CEK wrapped with RSA-OAEP-256) for the recipient, producing a JWE compact serialization.
/// </summary>
public class JoseEncryptHelper
{
    private readonly ITracingService _tracingService;
    private readonly KeyVaultKeyClient _keyVaultClient;
    private readonly string _signingKeyId;

    public JoseEncryptHelper(IServiceProvider serviceProvider)
    {
        _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);

        _signingKeyId = CommonIntegrationPlugin.GetEnvVariable(systemService, "ap_SigningKeyId");
        if (string.IsNullOrEmpty(_signingKeyId))
        {
            throw new InvalidOperationException("Signing key identifier (ap_SigningKeyId) is not configured.");
        }

        _keyVaultClient = new KeyVaultKeyClient(serviceProvider);
    }

    /// <summary>
    /// Signs <paramref name="plaintext"/> (ES256) then encrypts it (A256GCM, RSA-OAEP-256
    /// key wrap) for the recipient identified by <paramref name="recipientPublicKeyJwk"/>,
    /// returning the JWE compact serialization.
    /// </summary>
    public string Encrypt(string plaintext, string recipientPublicKeyJwk)
    {
        string jws = BuildSignedJws(plaintext);
        return EncryptJws(jws, recipientPublicKeyJwk);
    }

    private string BuildSignedJws(string plaintext)
    {
        string headerJson = "{\"alg\":\"ES256\"}";
        string header = Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(headerJson));
        string payload = Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(plaintext));
        string signingInput = header + "." + payload;

        // Key Vault's ES256 "sign" operation does not sign a hash of the raw
        // plaintext — it signs a SHA-256 hash of the JWS Signing Input, i.e.
        // SHA256(base64url(header) + "." + base64url(payload)), per RFC 7515. Hashing the
        // plaintext directly would produce a signature that fails verification against any
        // spec-compliant JWS validator.
        byte[] digest;
        using (var sha256 = SHA256.Create())
        {
            digest = sha256.ComputeHash(Encoding.ASCII.GetBytes(signingInput));
        }

        byte[] signature = _keyVaultClient.Sign(_signingKeyId, digest);
        string encodedSignature = Base64UrlEncoding.Encode(signature);

        return signingInput + "." + encodedSignature;
    }

    private string EncryptJws(string jws, string recipientPublicKeyJwk)
    {
        string jweHeaderJson = "{\"alg\":\"RSA-OAEP-256\",\"enc\":\"A256GCM\"}";
        string protectedHeader = Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(jweHeaderJson));

        // AAD for JWE compact serialization is the ASCII bytes of the base64url-encoded
        // protected header itself (RFC 7516 section 5.1).
        byte[] aad = Encoding.ASCII.GetBytes(protectedHeader);

        byte[] cek = new byte[32]; // AES-256 content encryption key
        byte[] iv = new byte[AesGcm256.NonceSizeBytes];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(cek);
            rng.GetBytes(iv);
        }

        byte[] plaintextBytes = Encoding.UTF8.GetBytes(jws);

        byte[] ciphertext;
        byte[] tag;
        AesGcm256.Encrypt(cek, iv, plaintextBytes, aad, out ciphertext, out tag);

        byte[] encryptedKey;
        using (RSA recipientRsaPublicKey = JwkHelper.RsaPublicKeyFromJwk(recipientPublicKeyJwk))
        {
            encryptedKey = recipientRsaPublicKey.Encrypt(cek, RSAEncryptionPadding.OaepSHA256);
        }

        // The CEK is a transient, in-memory-only value — never persisted.
        Array.Clear(cek, 0, cek.Length);

        return string.Join(".",
            protectedHeader,
            Base64UrlEncoding.Encode(encryptedKey),
            Base64UrlEncoding.Encode(iv),
            Base64UrlEncoding.Encode(ciphertext),
            Base64UrlEncoding.Encode(tag));
    }
}
