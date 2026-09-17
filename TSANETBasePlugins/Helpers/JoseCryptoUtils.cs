using System;
using System.Security.Cryptography;

/// <summary>
/// Base64url encode/decode helpers per RFC 7515 (JOSE) — standard base64 with
/// '+' -> '-', '/' -> '_', and no padding.
/// </summary>
internal static class Base64UrlEncoding
{
    public static string Encode(byte[] data)
    {
        return Convert.ToBase64String(data)
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
    }

    public static byte[] Decode(string value)
    {
        string s = value.Replace('-', '+').Replace('_', '/');
        switch (s.Length % 4)
        {
            case 2: s += "=="; break;
            case 3: s += "="; break;
        }
        return Convert.FromBase64String(s);
    }
}

/// <summary>
/// Minimal, self-contained AES-256-GCM implementation (per NIST SP 800-38D) built on
/// top of the framework's AES-ECB single-block primitive,
/// Only 96-bit (12-byte) nonces are supported, matching the JWE
/// A256GCM requirement (RFC 7518).
/// </summary>
internal static class AesGcm256
{
    private const int BlockSizeBytes = 16;
    public const int NonceSizeBytes = 12;
    public const int TagSizeBytes = 16;

    public static void Encrypt(byte[] key, byte[] nonce, byte[] plaintext, byte[] aad, out byte[] ciphertext, out byte[] tag)
    {
        if (nonce.Length != NonceSizeBytes)
        {
            throw new ArgumentException("Only 96-bit (12-byte) nonces are supported.", nameof(nonce));
        }

        using (Aes aes = CreateEcbAes(key))
        {
            byte[] h = EncryptBlock(aes, new byte[BlockSizeBytes]);
            byte[] j0 = ComputeJ0(nonce);
            ciphertext = Gctr(aes, IncrementCounter(j0), plaintext);
            byte[] s = Ghash(h, aad, ciphertext);
            byte[] encryptedJ0 = EncryptBlock(aes, j0);
            tag = Xor(s, encryptedJ0);
        }
    }

    public static byte[] Decrypt(byte[] key, byte[] nonce, byte[] ciphertext, byte[] tag, byte[] aad)
    {
        if (nonce.Length != NonceSizeBytes)
        {
            throw new ArgumentException("Only 96-bit (12-byte) nonces are supported.", nameof(nonce));
        }

        using (Aes aes = CreateEcbAes(key))
        {
            byte[] h = EncryptBlock(aes, new byte[BlockSizeBytes]);
            byte[] j0 = ComputeJ0(nonce);
            byte[] s = Ghash(h, aad, ciphertext);
            byte[] encryptedJ0 = EncryptBlock(aes, j0);
            byte[] expectedTag = Xor(s, encryptedJ0);

            if (!FixedTimeEquals(expectedTag, tag))
            {
                throw new CryptographicException("GCM authentication tag mismatch.");
            }

            return Gctr(aes, IncrementCounter(j0), ciphertext);
        }
    }

    private static byte[] ComputeJ0(byte[] nonce)
    {
        // For 96-bit nonces, J0 = nonce || 0x00000001 (32-bit counter starting at 1).
        byte[] j0 = new byte[BlockSizeBytes];
        Array.Copy(nonce, j0, NonceSizeBytes);
        j0[15] = 1;
        return j0;
    }

    private static Aes CreateEcbAes(byte[] key)
    {
        Aes aes = Aes.Create();
        aes.Key = key;
        aes.Mode = CipherMode.ECB;
        aes.Padding = PaddingMode.None;
        return aes;
    }

    private static byte[] EncryptBlock(Aes aes, byte[] block)
    {
        using (ICryptoTransform encryptor = aes.CreateEncryptor())
        {
            return encryptor.TransformFinalBlock(block, 0, block.Length);
        }
    }

    private static byte[] IncrementCounter(byte[] block)
    {
        byte[] result = (byte[])block.Clone();
        for (int i = 15; i >= 12; i--)
        {
            result[i]++;
            if (result[i] != 0)
            {
                break;
            }
        }
        return result;
    }

    private static byte[] Gctr(Aes aes, byte[] initialCounter, byte[] data)
    {
        if (data.Length == 0)
        {
            return new byte[0];
        }

        byte[] output = new byte[data.Length];
        byte[] counter = initialCounter;
        int offset = 0;
        while (offset < data.Length)
        {
            byte[] keystream = EncryptBlock(aes, counter);
            int len = Math.Min(BlockSizeBytes, data.Length - offset);
            for (int i = 0; i < len; i++)
            {
                output[offset + i] = (byte)(data[offset + i] ^ keystream[i]);
            }
            offset += BlockSizeBytes;
            counter = IncrementCounter(counter);
        }
        return output;
    }

    private static byte[] Ghash(byte[] h, byte[] aad, byte[] ciphertext)
    {
        byte[] y = new byte[BlockSizeBytes];
        y = GhashBlocks(y, h, aad);
        y = GhashBlocks(y, h, ciphertext);

        byte[] lengths = new byte[BlockSizeBytes];
        WriteUInt64BigEndian(lengths, 0, (ulong)aad.Length * 8);
        WriteUInt64BigEndian(lengths, 8, (ulong)ciphertext.Length * 8);
        for (int i = 0; i < BlockSizeBytes; i++)
        {
            y[i] ^= lengths[i];
        }

        return GaloisMultiply(y, h);
    }

    private static byte[] GhashBlocks(byte[] y, byte[] h, byte[] data)
    {
        int offset = 0;
        while (offset < data.Length)
        {
            byte[] block = new byte[BlockSizeBytes];
            int len = Math.Min(BlockSizeBytes, data.Length - offset);
            Array.Copy(data, offset, block, 0, len);
            for (int i = 0; i < BlockSizeBytes; i++)
            {
                y[i] ^= block[i];
            }
            y = GaloisMultiply(y, h);
            offset += BlockSizeBytes;
        }
        return y;
    }

    // GF(2^128) multiplication per NIST SP 800-38D, section 6.3 (bit-reflected / MSB-first
    // representation, reduction polynomial R = 0xE1 followed by 120 zero bits).
    private static byte[] GaloisMultiply(byte[] x, byte[] y)
    {
        byte[] z = new byte[BlockSizeBytes];
        byte[] v = (byte[])y.Clone();

        for (int i = 0; i < 128; i++)
        {
            int byteIndex = i / 8;
            int bitIndex = 7 - (i % 8);
            bool xBitSet = (x[byteIndex] & (1 << bitIndex)) != 0;

            if (xBitSet)
            {
                for (int k = 0; k < BlockSizeBytes; k++)
                {
                    z[k] ^= v[k];
                }
            }

            bool lsbSet = (v[15] & 1) != 0;
            for (int k = 15; k > 0; k--)
            {
                v[k] = (byte)((v[k] >> 1) | ((v[k - 1] & 1) << 7));
            }
            v[0] = (byte)(v[0] >> 1);

            if (lsbSet)
            {
                v[0] ^= 0xE1;
            }
        }

        return z;
    }

    private static byte[] Xor(byte[] a, byte[] b)
    {
        byte[] result = new byte[a.Length];
        for (int i = 0; i < a.Length; i++)
        {
            result[i] = (byte)(a[i] ^ b[i]);
        }
        return result;
    }

    private static bool FixedTimeEquals(byte[] a, byte[] b)
    {
        if (a.Length != b.Length)
        {
            return false;
        }

        int diff = 0;
        for (int i = 0; i < a.Length; i++)
        {
            diff |= a[i] ^ b[i];
        }
        return diff == 0;
    }

    private static void WriteUInt64BigEndian(byte[] buffer, int offset, ulong value)
    {
        for (int i = 7; i >= 0; i--)
        {
            buffer[offset + i] = (byte)(value & 0xFF);
            value >>= 8;
        }
    }
}
