using System.Collections.Generic;

namespace TSANETBasePlugins.E2ee
{
    public class DecryptionContext
    {
        /// <summary>EC P-256 public JWK of the expected sender (must contain "kid").</summary>
        public string SenderSigningKeyJwk { get; set; }
        public string OwnCompanyId { get; set; }
        public string CounterpartCompanyId { get; set; }
        public string ExpectedRef { get; set; }
        public string ExpectedCid { get; set; }
        public bool E2eeRequired { get; set; }
    }

    public class DecryptResult
    {
        public bool IsVerified { get; set; }
        public DecryptStatus Status { get; set; }

        /// <summary>Plaintext payload; null unless verified.</summary>
        public string PayloadJson { get; set; }
        public string Jti { get; set; }
        public long? Iat { get; set; }
        public string ManifestDigest { get; set; }
    }

    public class ManifestFieldClaim
    {
        public string Name { get; set; }
        public string Alg { get; set; }
        public string Dig { get; set; }
    }

    /// <summary>Claims of a manifest. Unless the manifest signature was verified, never trust these.</summary>
    public class ManifestClaims
    {
        public string Ver { get; set; }
        public string Iss { get; set; }
        public string Aud { get; set; }
        public string Jti { get; set; }
        public long Iat { get; set; }
        public string Ref { get; set; }
        public string Cid { get; set; }
        public List<ManifestFieldClaim> Fields { get; set; } = new List<ManifestFieldClaim>();
    }
}
