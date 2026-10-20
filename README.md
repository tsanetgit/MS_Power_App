Microsoft Power App to be used in MS Dynamics and MS Teams
For access to the Beta and Production environments, contact membership@tsanet.org

## GitBook Documentation
https://tsanet.gitbook.io/connect/documentation/dynamics-app/demo

## Install Guide

The install guide provides information on how to install the Power App and details about the application and integration examples:
New guide for the Dynamics on the App Store: 
https://github.com/tsanetgit/MS_Power_App/blob/main/Connect_2.0_Dynamics_Package_Install_Guide_2_4.docx

Previous guide for the Power App manual installation using packages:
https://github.com/tsanetgit/MS_Power_App/blob/main/Connect%202.0%20Power%20App%20Package%20Install%20Guide%20v1_1.docx

## User Guide

The User Guide provides details on how to use the Power App 
https://github.com/tsanetgit/MS_Power_App/blob/master/Microsoft%20Power%20App%20-%20User%20Guide.docx

## Note encryption

Collaboration requests and notes use the same field-level E2EE profile. Outgoing
notes encrypt `summary` and `description` only when an `encryptedRecipient`
setting exists, just like outgoing collaboration requests. Each note carries its
own signed `manifest` and `encrypted` flag. Missing or invalid encryption
configuration fails the operation; it does not fall back to plaintext.

Notes are bound to their parent collaboration request's `submitterCaseNumber`
(`ap_submittercasenumber` in Dataverse), not their note ID or case token. The
note action and note-create plugin prepare and encrypt notes using the shared
E2EE helpers, then verify/decrypt successful standalone note responses.
`CommonIntegrationPlugin.PostCaseNote` only posts the prepared JSON and returns
the API response; it does not resolve the parent or perform encryption/decryption.
Every item in a collaboration response's `caseNotes` array is independently
verified and decrypted before being returned or saved.
Encrypted notes are processed even when the parent is plaintext; mixed arrays
and notes from either participant are supported. Plaintext notes remain unchanged.
A missing/invalid manifest or failed verification rejects the response before
case/note persistence. Decrypted payload content is not written to tracing.

Configuration reads use the existing SYSTEM-context service. Encryption still
uses the existing `encryptedRecipient`, `Iss`, signing-key and unwrap-key
configuration.

## Provide Feedback

Log an issue to report a bug or request changes
