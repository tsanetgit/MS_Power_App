using System;

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
