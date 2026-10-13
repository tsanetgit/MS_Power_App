namespace TSANETBasePlugins.E2ee
{
    /// <summary>Kind of payload whose fields are protected.</summary>
    public enum PayloadType
    {
        CollaborationRequest,
        Note
    }

    /// <summary>Outcome of a decrypt/verify operation.</summary>
    public enum DecryptStatus
    {
        Ok,
        NotEncrypted,
        ManifestMissing,
        SignatureInvalid,
        BindingMismatch,
        FieldSetMismatch,
        FingerprintMismatch,
        NoMatchingRecipient,
        AeadFailure,
        HeaderBindingMismatch
    }
}
