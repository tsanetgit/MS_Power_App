namespace TSANETBasePlugins.E2ee
{
    public class EncryptionContext
    {
        /// <summary>Recipient RSA public JWK; must contain "kid".</summary>
        public string RecipientEncryptionKeyJwk { get; set; }

        /// <summary>Own RSA public JWK (self-wrap); must contain "kid".</summary>
        public string OwnEncryptionKeyJwk { get; set; }

        /// <summary>Own TSANet company id.</summary>
        public string Iss { get; set; }

        /// <summary>Counterpart TSANet company id.</summary>
        public string Aud { get; set; }

        /// <summary>submitterCaseNumber (for notes: the parent case's value).</summary>
        public string Ref { get; set; }

        /// <summary>Optional; absent on the initial POST.</summary>
        public string Cid { get; set; }
    }

    public class EncryptResult
    {
        public string PayloadJson { get; set; }
        public string Jti { get; set; }
        public string Manifest { get; set; }
        public string ManifestDigest { get; set; }
    }
}
