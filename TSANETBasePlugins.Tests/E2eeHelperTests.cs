extern alias PluginAsm;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using PluginAsm::TSANETBasePlugins.E2ee;

namespace TSANETBasePlugins.Tests
{
    [TestClass]
    public class E2eeHelperTests
    {
        private const string CaseJson =
            "{\"submitterCaseNumber\":\"S-1\",\"problemSummary\":\"Sum secret\",\"problemDescription\":\"Desc secret\"," +
            "\"priority\":\"HIGH\",\"amount\":1.10,\"opened\":\"2024-01-01T00:00:00Z\",\"nothing\":null," +
            "\"customFields\":[" +
            "{\"fieldId\":11,\"label\":\"A\",\"value\":\"cf-one\",\"selections\":[{\"value\":\"keep\",\"children\":[\"c\"]}]}," +
            "{\"fieldId\":12,\"label\":\"B\",\"value\":\"cf-two\"}," +
            "{\"fieldId\":13,\"label\":\"C\",\"value\":null}]}";

        private const string NoteJson = "{\"summary\":\"Note sum\",\"description\":\"Note desc\",\"createdBy\":\"x\",\"seq\":5}";

        private Party _alice;
        private Party _bob;

        [TestInitialize]
        public void Init()
        {
            _alice = new Party("alice");
            _bob = new Party("bob");
        }

        private EncryptionContext EncCtx(string cid = null)
        {
            return new EncryptionContext
            {
                RecipientEncryptionKeyJwk = _bob.EncJwk(),
                OwnEncryptionKeyJwk = _alice.EncJwk(),
                Iss = "ALICE",
                Aud = "BOB",
                Ref = "S-1",
                Cid = cid
            };
        }

        private DecryptionContext BobCtx()
        {
            return new DecryptionContext
            {
                SenderSigningKeyJwk = _alice.SigningJwk(),
                OwnCompanyId = "BOB",
                CounterpartCompanyId = "ALICE",
                ExpectedRef = "S-1",
                E2eeRequired = true
            };
        }

        private DecryptionContext AliceCtx()
        {
            DecryptionContext c = BobCtx();
            c.OwnCompanyId = "ALICE";
            c.CounterpartCompanyId = "BOB";
            return c;
        }

        /// <summary>The API returns custom fields with "id" instead of the submitted "fieldId".</summary>
        private static string ToIncomingShape(string json)
        {
            JObject o = JObject.Parse(json);
            var cf = o["customFields"] as JArray;
            if (cf != null)
            {
                foreach (JObject item in cf)
                {
                    JProperty p = item.Property("fieldId");
                    p.Replace(new JProperty("id", (int)p.Value + 10000));
                }
            }
            RenameProperty(o, "problemSummary", "summary");
            RenameProperty(o, "problemDescription", "description");
            JToken desc = o["description"];
            if (desc != null && desc.Type == JTokenType.String)
            {
                o["description"] = ((string)desc).Replace("\"", "&#34;");
            }
            return o.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static void RenameProperty(JObject o, string from, string to)
        {
            JProperty p = o.Property(from);
            if (p != null)
            {
                p.Replace(new JProperty(to, p.Value));
            }
        }

        private EncryptResult EncryptCase(string cid = null)
        {
            return _alice.Encryptor().Encrypt(PayloadType.CollaborationRequest, CaseJson, EncCtx(cid));
        }

        // ---- round trips ----

        [TestMethod]
        public void RoundTrip_CollaborationRequest_MultipleCustomFields()
        {
            EncryptResult enc = EncryptCase();
            JObject encrypted = JObject.Parse(enc.PayloadJson);
            Assert.IsTrue((bool)encrypted["encrypted"]);
            Assert.AreEqual(enc.Manifest, (string)encrypted["manifest"]);
            Assert.IsFalse(enc.PayloadJson.Contains("Sum secret"));
            Assert.IsFalse(enc.PayloadJson.Contains("cf-one"));

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), BobCtx());

            Assert.AreEqual(DecryptStatus.Ok, dec.Status);
            Assert.IsTrue(dec.IsVerified);
            Assert.AreEqual(enc.Jti, dec.Jti);
            Assert.AreEqual(enc.ManifestDigest, dec.ManifestDigest);
            Assert.IsTrue(dec.Iat.HasValue);
            JObject plain = JObject.Parse(dec.PayloadJson);
            Assert.AreEqual("Sum secret", (string)plain["summary"]);
            Assert.AreEqual("Desc secret", (string)plain["description"]);
            Assert.AreEqual("cf-one", (string)plain["customFields"][0]["value"]);
            Assert.AreEqual("cf-two", (string)plain["customFields"][1]["value"]);
            Assert.AreEqual(JTokenType.Null, plain["customFields"][2]["value"].Type);
        }

        [TestMethod]
        public void RoundTrip_Note()
        {
            EncryptResult enc = _alice.Encryptor().Encrypt(PayloadType.Note, NoteJson, EncCtx());
            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.Note, enc.PayloadJson, BobCtx());

            Assert.AreEqual(DecryptStatus.Ok, dec.Status);
            JObject plain = JObject.Parse(dec.PayloadJson);
            Assert.AreEqual("Note sum", (string)plain["summary"]);
            Assert.AreEqual("Note desc", (string)plain["description"]);
        }

        [TestMethod]
        public void Manifest_ListsExpectedFieldIdentifiers_NeverArrayIndex()
        {
            EncryptResult enc = EncryptCase(cid: "CID-9");
            ManifestClaims claims = _bob.Decryptor().ReadClaimsUnverified(enc.PayloadJson);

            CollectionAssert.AreEquivalent(
                new[] { "problemSummary", "problemDescription", "customFields/11/value", "customFields/12/value" },
                claims.Fields.Select(f => f.Name).ToArray());
            Assert.AreEqual("tsanet-e2ee-1", claims.Ver);
            Assert.AreEqual("CID-9", claims.Cid);
            Assert.AreEqual("S-1", claims.Ref);
            Assert.IsTrue(claims.Fields.All(f => f.Alg == "S256"));
        }

        [TestMethod]
        public void Encrypt_LeavesOtherPropertiesByteIdentical()
        {
            EncryptResult enc = EncryptCase();

            StringAssert.Contains(enc.PayloadJson, "\"amount\":1.10");
            StringAssert.Contains(enc.PayloadJson, "\"opened\":\"2024-01-01T00:00:00Z\"");
            StringAssert.Contains(enc.PayloadJson, "\"nothing\":null");
            StringAssert.Contains(enc.PayloadJson, "\"selections\":[{\"value\":\"keep\",\"children\":[\"c\"]}]");
            StringAssert.Contains(enc.PayloadJson, "\"priority\":\"HIGH\"");

            JObject encrypted = JObject.Parse(enc.PayloadJson);
            Assert.AreEqual(JTokenType.Null, encrypted["customFields"][2]["value"].Type);
        }

        [TestMethod]
        public void Sender_CanDecryptOwnPayload_ViaSelfWrap()
        {
            EncryptResult enc = EncryptCase();
            DecryptResult dec = _alice.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), AliceCtx());

            Assert.AreEqual(DecryptStatus.Ok, dec.Status);
            Assert.AreEqual("Sum secret", (string)JObject.Parse(dec.PayloadJson)["summary"]);
        }

        [TestMethod]
        public void Traces_NeverContainPlaintext()
        {
            EncryptResult enc = EncryptCase();
            _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), BobCtx());

            string all = string.Join("\n", _alice.Trace.Lines.Concat(_bob.Trace.Lines));
            Assert.IsTrue(all.Length > 0);
            Assert.IsFalse(all.Contains("Sum secret"));
            Assert.IsFalse(all.Contains("cf-one"));
        }

        // ---- tampering ----

        [TestMethod]
        public void Tamper_ModifiedCiphertext_FingerprintMismatch()
        {
            EncryptResult enc = EncryptCase();
            JObject p = JObject.Parse(ToIncomingShape(enc.PayloadJson));
            p["summary"] = FlipCiphertext((string)p["summary"]);

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.FingerprintMismatch);
        }

        [TestMethod]
        public void Tamper_ModifiedCiphertext_WithResignedManifest_AeadFailure()
        {
            EncryptResult enc = EncryptCase();
            JObject p = JObject.Parse(ToIncomingShape(enc.PayloadJson));
            string tampered = FlipCiphertext((string)p["summary"]);
            p["summary"] = tampered;
            p["manifest"] = _alice.Forge((string)p["manifest"], c => SetDig(c, "problemSummary", tampered));

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.AeadFailure);
        }

        [TestMethod]
        public void Tamper_FieldSwappedBetweenPayloads_FingerprintMismatch()
        {
            JObject a = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            JObject b = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            a["summary"] = b["summary"];

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, a.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.FingerprintMismatch);
        }

        [TestMethod]
        public void Tamper_FieldSwappedBetweenPayloads_WithResignedManifest_HeaderBindingMismatch()
        {
            JObject a = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            JObject b = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            string foreign = (string)b["summary"];
            a["summary"] = foreign;
            a["manifest"] = _alice.Forge((string)a["manifest"], c => SetDig(c, "problemSummary", foreign));

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, a.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.HeaderBindingMismatch);
        }

        [TestMethod]
        public void Tamper_FieldMovedToOtherFieldName_WithResignedManifest_HeaderBindingMismatch()
        {
            JObject a = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            string summaryJwe = (string)a["summary"];
            a["description"] = summaryJwe;
            a["manifest"] = _alice.Forge((string)a["manifest"], c => SetDig(c, "problemDescription", summaryJwe));

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, a.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.HeaderBindingMismatch);
        }

        [TestMethod]
        public void Tamper_RemovedField_FieldSetMismatch()
        {
            JObject p = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            p.Remove("description");

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx());

            AssertFailed(dec, DecryptStatus.FieldSetMismatch);
        }

        [TestMethod]
        public void Tamper_ChangedRef_BindingMismatch()
        {
            EncryptResult enc = EncryptCase();
            DecryptionContext ctx = BobCtx();
            ctx.ExpectedRef = "S-OTHER";

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), ctx), DecryptStatus.BindingMismatch);
        }

        [TestMethod]
        public void Tamper_ChangedRefInManifest_WithoutResign_SignatureInvalid()
        {
            JObject p = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            p["manifest"] = EditClaimsWithoutResign((string)p["manifest"], c => c["ref"] = "S-EVIL");

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx()), DecryptStatus.SignatureInvalid);
        }

        [TestMethod]
        public void Tamper_ChangedIss_WithResignedManifest_BindingMismatch()
        {
            JObject p = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            p["manifest"] = _alice.Forge((string)p["manifest"], c => c["iss"] = "MALLORY");

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx()), DecryptStatus.BindingMismatch);
        }

        [TestMethod]
        public void Tamper_ChangedIss_WithoutResign_SignatureInvalid()
        {
            JObject p = JObject.Parse(ToIncomingShape(EncryptCase().PayloadJson));
            p["manifest"] = EditClaimsWithoutResign((string)p["manifest"], c => c["iss"] = "MALLORY");

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, p.ToString(), BobCtx()), DecryptStatus.SignatureInvalid);
        }

        [TestMethod]
        public void Tamper_CidMismatch_BindingMismatch()
        {
            EncryptResult enc = EncryptCase(cid: "CID-1");
            DecryptionContext ctx = BobCtx();
            ctx.ExpectedCid = "CID-2";

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), ctx), DecryptStatus.BindingMismatch);
        }

        [TestMethod]
        public void WrongSigningKey_SignatureInvalid()
        {
            EncryptResult enc = EncryptCase();
            var mallory = new Party("mallory");
            DecryptionContext ctx = BobCtx();
            ctx.SenderSigningKeyJwk = mallory.SigningJwk(_alice.SigningKid);

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), ctx), DecryptStatus.SignatureInvalid);
        }

        [TestMethod]
        public void SenderKidDiffersFromManifestKid_SignatureInvalid()
        {
            EncryptResult enc = EncryptCase();
            DecryptionContext ctx = BobCtx();
            ctx.SenderSigningKeyJwk = _alice.SigningJwk("https://alice.vault.azure.net/keys/sign/other");

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), ctx), DecryptStatus.SignatureInvalid);
        }

        [TestMethod]
        public void UnknownRecipientKid_NoMatchingRecipient()
        {
            EncryptResult enc = EncryptCase();
            var carol = new Party("carol");

            DecryptResult dec = carol.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), new DecryptionContext
            {
                SenderSigningKeyJwk = _alice.SigningJwk(),
                OwnCompanyId = "BOB",
                CounterpartCompanyId = "ALICE",
                ExpectedRef = "S-1",
                E2eeRequired = true
            });

            AssertFailed(dec, DecryptStatus.NoMatchingRecipient);
            Assert.AreEqual(0, carol.Vault.UnwrapCalls);
        }

        [TestMethod]
        public void RecipientKidOutsideVault_NotUsed()
        {
            EncryptResult enc = EncryptCase();
            Party bobWithOtherVault = new Party("bob");
            bobWithOtherVault.Config.KeyVaultUri = "https://elsewhere.vault.azure.net";
            bobWithOtherVault.Config.EncryptionKeyId = "https://elsewhere.vault.azure.net/keys/enc/ev1";

            DecryptResult dec = bobWithOtherVault.Decryptor().Decrypt(PayloadType.CollaborationRequest, ToIncomingShape(enc.PayloadJson), BobCtx());

            AssertFailed(dec, DecryptStatus.NoMatchingRecipient);
        }

        // ---- missing manifest ----

        [TestMethod]
        public void MissingManifest_Required_ManifestMissing()
        {
            DecryptionContext ctx = BobCtx();
            ctx.E2eeRequired = true;

            AssertFailed(_bob.Decryptor().Decrypt(PayloadType.Note, NoteJson, ctx), DecryptStatus.ManifestMissing);
        }

        [TestMethod]
        public void MissingManifest_NotRequired_ReturnedUnchanged()
        {
            DecryptionContext ctx = BobCtx();
            ctx.E2eeRequired = false;

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.Note, NoteJson, ctx);

            Assert.IsTrue(dec.IsVerified);
            Assert.AreEqual(DecryptStatus.NotEncrypted, dec.Status);
            Assert.AreEqual(NoteJson, dec.PayloadJson);
        }

        [TestMethod]
        public void EncryptedFlagIsAdvisory_IgnoredWithoutManifest()
        {
            DecryptionContext ctx = BobCtx();
            ctx.E2eeRequired = false;
            string json = "{\"summary\":\"s\",\"description\":\"d\",\"encrypted\":true}";

            DecryptResult dec = _bob.Decryptor().Decrypt(PayloadType.Note, json, ctx);

            Assert.AreEqual(DecryptStatus.NotEncrypted, dec.Status);
        }

        // ---- genuine faults ----

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException))]
        public void MalformedPayload_Throws()
        {
            _bob.Decryptor().Decrypt(PayloadType.Note, "{not json", BobCtx());
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException))]
        public void NonStringCustomFieldValue_Throws()
        {
            string json = "{\"problemSummary\":\"a\",\"customFields\":[{\"fieldId\":1,\"value\":5}]}";
            _alice.Encryptor().Encrypt(PayloadType.CollaborationRequest, json, EncCtx());
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException))]
        public void MissingConfiguration_Throws()
        {
            var enc = new E2eePayloadEncryptor(_alice.Vault, new E2eeConfig(), _alice.Trace);
            enc.Encrypt(PayloadType.Note, NoteJson, EncCtx());
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException))]
        public void RecipientJwkWithoutKid_Throws()
        {
            EncryptionContext ctx = EncCtx();
            JObject jwk = JObject.Parse(ctx.RecipientEncryptionKeyJwk);
            jwk.Remove("kid");
            ctx.RecipientEncryptionKeyJwk = jwk.ToString();
            _alice.Encryptor().Encrypt(PayloadType.Note, NoteJson, ctx);
        }

        // ---- test vectors ----

        [TestMethod]
        public void EmitTestVectors()
        {
            EncryptResult enc = _alice.Encryptor().Encrypt(PayloadType.Note, NoteJson, EncCtx(cid: "CID-1"));
            JObject payload = JObject.Parse(enc.PayloadJson);
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestVectors");
            Directory.CreateDirectory(dir);

            File.WriteAllText(Path.Combine(dir, "manifest.jws"), enc.Manifest);
            File.WriteAllText(Path.Combine(dir, "two-recipient-jwe.json"), (string)payload["summary"]);
            File.WriteAllText(Path.Combine(dir, "vector-info.json"), new JObject
            {
                ["profile"] = "tsanet-e2ee-1",
                ["note"] = "TEST ONLY keys generated for this vector. Field 'summary' of payload; plaintext = 'Note sum'.",
                ["jti"] = enc.Jti,
                ["manifestDigest"] = enc.ManifestDigest,
                ["senderSigningPublicJwk"] = JObject.Parse(_alice.SigningJwk()),
                ["recipientPublicJwk"] = JObject.Parse(_bob.EncJwk()),
                ["senderPublicJwk"] = JObject.Parse(_alice.EncJwk()),
                ["iss"] = "ALICE",
                ["aud"] = "BOB",
                ["ref"] = "S-1",
                ["fullPayload"] = payload
            }.ToString());

            Assert.AreEqual(2, JObject.Parse((string)payload["summary"])["recipients"].Count());
        }

        // ---- helpers ----

        private static void AssertFailed(DecryptResult result, DecryptStatus expected)
        {
            Assert.AreEqual(expected, result.Status);
            Assert.IsFalse(result.IsVerified);
            Assert.IsNull(result.PayloadJson);
        }

        private static string FlipCiphertext(string jwe)
        {
            JObject o = JObject.Parse(jwe);
            byte[] ct = B64.Decode((string)o["ciphertext"]);
            ct[0] ^= 0x01;
            o["ciphertext"] = B64.Encode(ct);
            return o.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static void SetDig(JObject claims, string fieldName, string fieldValue)
        {
            foreach (JObject f in (JArray)claims["fields"])
            {
                if ((string)f["name"] == fieldName)
                {
                    f["dig"] = B64.Sha256(fieldValue);
                }
            }
        }

        private static string EditClaimsWithoutResign(string manifest, Action<JObject> mutate)
        {
            string[] parts = manifest.Split('.');
            JObject claims = JObject.Parse(System.Text.Encoding.UTF8.GetString(B64.Decode(parts[1])));
            mutate(claims);
            parts[1] = B64.Encode(System.Text.Encoding.UTF8.GetBytes(claims.ToString(Newtonsoft.Json.Formatting.None)));
            return string.Join(".", parts);
        }
    }
}

