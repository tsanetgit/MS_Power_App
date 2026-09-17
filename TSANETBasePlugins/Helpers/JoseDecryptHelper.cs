using Microsoft.Xrm.Sdk;
using System;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// Implements the "decrypt-then-verify" the JWE compact serialization is decrypted (CEK recovered via Key Vault's
/// RSA-OAEP-256 "unwrapKey" operation, payload decrypted locally with A256GCM) to recover
/// the inner JWS, which is then verified (ES256) against the sender's public key.
/// </summary>
public class JoseDecryptHelper
{
    private readonly ITracingService _tracingService;
    private readonly KeyVaultKeyClient _keyVaultClient;
    private readonly string _unwrapKeyId;

    public JoseDecryptHelper(IServiceProvider serviceProvider)
    {
        _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);

        _unwrapKeyId = CommonIntegrationPlugin.GetEnvVariable(systemService, "ap_UnwrapKeyId");
        if (string.IsNullOrEmpty(_unwrapKeyId))
        {
            throw new InvalidOperationException("Unwrap key identifier (ap_UnwrapKeyId) is not configured.");
        }

        _keyVaultClient = new KeyVaultKeyClient(serviceProvider);
    }

    /// <summary>
    /// Decrypts <paramref name="jweCompact"/> and verifies the inner JWS (ES256) against
    /// <paramref name="senderPublicKeyJwk"/>. On successful verification, returns the
    /// decrypted plaintext with <c>verified = true</c>. On verification failure, the
    /// unverified plaintext is withheld — <c>plaintext</c> is returned as
    /// <see langword="null"/> and <c>verified</c> is <see langword="false"/>, so the caller
    /// can quarantine the message per the E2E encryption proposal's failure-handling
    /// requirement. Genuine faults (malformed JWE, Key Vault errors) still throw.
    /// </summary>
    public (string plaintext, bool verified) Decrypt(string jweCompact, string senderPublicKeyJwk)
    {
        string jws = DecryptJwe(jweCompact);
        return VerifyJws(jws, senderPublicKeyJwk);
    }

    private string DecryptJwe(string jweCompact)
    {
        string[] parts = jweCompact.Split('.');
        if (parts.Length != 5)
        {
            throw new InvalidOperationException("Malformed JWE compact serialization: expected 5 dot-separated segments.");
        }

        string protectedHeader = parts[0];
        byte[] encryptedKey = Base64UrlEncoding.Decode(parts[1]);
        byte[] iv = Base64UrlEncoding.Decode(parts[2]);
        byte[] ciphertext = Base64UrlEncoding.Decode(parts[3]);
        byte[] tag = Base64UrlEncoding.Decode(parts[4]);

        // AAD for JWE compact serialization is the ASCII bytes of the base64url-encoded
        // protected header itself (RFC 7516 section 5.2).
        byte[] aad = Encoding.ASCII.GetBytes(protectedHeader);

        byte[] cek = _keyVaultClient.UnwrapKey(_unwrapKeyId, encryptedKey);
        try
        {
            byte[] plaintextBytes = AesGcm256.Decrypt(cek, iv, ciphertext, tag, aad);
            return Encoding.UTF8.GetString(plaintextBytes);
        }
        finally
        {
            // The CEK is a transient, in-memory-only value — never persisted.
            Array.Clear(cek, 0, cek.Length);
        }
    }

    private (string plaintext, bool verified) VerifyJws(string jws, string senderPublicKeyJwk)
    {
        string[] parts = jws.Split('.');
        if (parts.Length != 3)
        {
            throw new InvalidOperationException("Malformed JWS: expected 3 dot-separated segments.");
        }

        string header = parts[0];
        string payload = parts[1];
        string signingInput = header + "." + payload;

        byte[] digest;
        using (var sha256 = SHA256.Create())
        {
            digest = sha256.ComputeHash(Encoding.ASCII.GetBytes(signingInput));
        }

        byte[] signature;
        try
        {
            signature = Base64UrlEncoding.Decode(parts[2]);
        }
        catch (FormatException ex)
        {
            _tracingService?.Trace($"JWS signature segment is not valid base64url: {ex.Message}");
            return (null, false);
        }

        bool verified;
        try
        {
            using (ECDsa senderPublicKey = JwkHelper.EcPublicKeyFromJwk(senderPublicKeyJwk))
            {
                verified = senderPublicKey.VerifyHash(digest, signature);
            }
        }
        catch (CryptographicException ex)
        {
            _tracingService?.Trace($"ES256 signature verification failed: {ex.Message}");
            verified = false;
        }

        if (!verified)
        {
            _tracingService?.Trace("JWS signature verification failed — withholding unverified plaintext.");
            return (null, false);
        }

        string plaintext = Encoding.UTF8.GetString(Base64UrlEncoding.Decode(payload));
        return (plaintext, true);
    }
}
