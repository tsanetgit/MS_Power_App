extern alias PluginAsm;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using PluginAsm::TSANETBasePlugins.E2ee;

namespace TSANETBasePlugins.Tests
{
    internal static class B64
    {
        public static string Encode(byte[] data)
        {
            return Convert.ToBase64String(data).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }

        public static byte[] Decode(string s)
        {
            s = s.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            return Convert.FromBase64String(s);
        }

        public static string Sha256(string utf8)
        {
            using (var sha = SHA256.Create())
            {
                return Encode(sha.ComputeHash(Encoding.UTF8.GetBytes(utf8)));
            }
        }
    }

    internal class TraceRecorder : ITracingService
    {
        public readonly List<string> Lines = new List<string>();

        public void Trace(string format, params object[] args)
        {
            Lines.Add(args == null || args.Length == 0 ? format : string.Format(format, args));
        }
    }

    /// <summary>In-memory stand-in for Key Vault: one EC signing key and any number of RSA keys.</summary>
    internal class FakeKeyVault : IKeyVaultCrypto
    {
        public ECDsaCng SigningKey;
        public Dictionary<string, RSACng> RsaKeys = new Dictionary<string, RSACng>();
        public int UnwrapCalls;

        public byte[] Sign(byte[] digest)
        {
            return SigningKey.SignHash(digest);
        }

        public byte[] Unwrap(string kid, byte[] encryptedKey)
        {
            UnwrapCalls++;
            RSACng key;
            if (!RsaKeys.TryGetValue(kid, out key))
            {
                throw new InvalidPluginExecutionException("Key not found: " + kid);
            }
            return key.Decrypt(encryptedKey, RSAEncryptionPadding.OaepSHA256);
        }
    }

    /// <summary>A TSANet member with its own vault, signing key and encryption key.</summary>
    internal class Party
    {
        public string Name;
        public string SigningKid;
        public string EncKid;
        public ECDsaCng Signing = new ECDsaCng(256);
        public RSACng Enc = new RSACng(2048);
        public FakeKeyVault Vault = new FakeKeyVault();
        public E2eeConfig Config;
        public TraceRecorder Trace = new TraceRecorder();

        public Party(string name)
        {
            Name = name;
            string vault = "https://" + name + ".vault.azure.net";
            SigningKid = vault + "/keys/sign/sv1";
            EncKid = vault + "/keys/enc/ev1";
            Vault.SigningKey = Signing;
            Vault.RsaKeys[EncKid] = Enc;
            Config = new E2eeConfig { KeyVaultUri = vault, SigningKeyId = SigningKid, EncryptionKeyId = EncKid };
        }

        public E2eePayloadEncryptor Encryptor() { return new E2eePayloadEncryptor(Vault, Config, Trace); }
        public E2eePayloadDecryptor Decryptor() { return new E2eePayloadDecryptor(Vault, Config, Trace); }

        public string SigningJwk(string kidOverride = null)
        {
            byte[] blob = Signing.Key.Export(CngKeyBlobFormat.EccPublicBlob);
            byte[] x = new byte[32], y = new byte[32];
            Array.Copy(blob, 8, x, 0, 32);
            Array.Copy(blob, 40, y, 0, 32);
            return new JObject { ["kty"] = "EC", ["crv"] = "P-256", ["kid"] = kidOverride ?? SigningKid, ["x"] = B64.Encode(x), ["y"] = B64.Encode(y) }.ToString();
        }

        public string EncJwk()
        {
            RSAParameters p = Enc.ExportParameters(false);
            return new JObject { ["kty"] = "RSA", ["kid"] = EncKid, ["n"] = B64.Encode(p.Modulus), ["e"] = B64.Encode(p.Exponent) }.ToString();
        }

        /// <summary>Re-signs a manifest after mutating its claims (simulates an attacker holding the signing key).</summary>
        public string Forge(string manifest, Action<JObject> mutateClaims)
        {
            string[] parts = manifest.Split('.');
            JObject claims = JObject.Parse(Encoding.UTF8.GetString(B64.Decode(parts[1])));
            mutateClaims(claims);
            string input = parts[0] + "." + B64.Encode(Encoding.UTF8.GetBytes(claims.ToString(Newtonsoft.Json.Formatting.None)));
            using (var sha = SHA256.Create())
            {
                return input + "." + B64.Encode(Signing.SignHash(sha.ComputeHash(Encoding.ASCII.GetBytes(input))));
            }
        }
    }
}

