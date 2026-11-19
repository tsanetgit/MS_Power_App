extern alias PluginAsm;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using PluginAsm::TSANETBasePlugins.E2ee;

namespace TSANETBasePlugins.Tests
{
    [TestClass]
    public class E2eeRecipientTests
    {
        [TestMethod]
        public void NewCase_UnconfiguredRecipientRemainsPlaintext()
        {
            var service = new SettingsService();
            service.Add("encryptedRecipient", "11", "https://example.test/keys");
            string json = "{\"receiveCompanyId\":11}";
            string result = E2eeCaseEncryption.PrepareForPost(new Services(service), PayloadType.CollaborationRequest, json, null, "42");
            Assert.AreEqual(json, result);
            CollectionAssert.AreEqual(new[] { "encryptedRecipient", "42" }, service.RecipientConditions);
        }

        [TestMethod]
        public void NewCase_PluginUsesStringActionRecipientRatherThanPayloadRecipient()
        {
            var parameters = new ParameterCollection
            {
                { "CaseDetails", "{\"receiveCompanyId\":11}" },
                { "CaseID", "" },
                { "RecipientCompanyID", "42" }
            };
            var service = new SettingsService();
            service.Add("encryptedRecipient", "11", null);
            service.Add("encryptedRecipient", "42", "https://example.test/keys");
            var context = (IPluginExecutionContext)new PluginContextProxy(parameters).GetTransparentProxy();
            var ex = Assert.ThrowsException<InvalidPluginExecutionException>(() =>
                new PluginAsm::PostCasePlugin().Execute(new Services(service, context)));
            StringAssert.Contains(ex.Message, "Iss setting (ap_value)");
            CollectionAssert.AreEqual(new[] { "encryptedRecipient", "42" }, service.RecipientConditions);
        }

        [DataTestMethod]
        [DataRow(0, "11")]
        [DataRow(1, "22")]
        public void Note_UsesCounterpartyFromParentResponse(int direction, string expected)
        {
            var parent = new Entity("ap_tsanetcase");
            parent["ap_direction"] = new OptionSetValue(direction);
            parent["ap_companycode"] = 999;
            parent["ap_formjson"] = "{\"submitCompanyId\":11,\"receiveCompanyId\":22}";
            Assert.AreEqual(expected, E2eeCaseEncryption.GetNoteRecipientCompanyId(parent));
        }

        [TestMethod]
        public void MissingRecipientOrParentDirection_FailsInsteadOfGuessing()
        {
            var parent = new Entity("ap_tsanetcase");
            Assert.ThrowsException<InvalidPluginExecutionException>(() => E2eeCaseEncryption.GetNoteRecipientCompanyId(parent));
            parent["ap_direction"] = new OptionSetValue(1);
            parent["ap_formjson"] = "{\"submitCompanyId\":11}";
            Assert.ThrowsException<InvalidPluginExecutionException>(() => E2eeCaseEncryption.GetNoteRecipientCompanyId(parent));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow(" ")]
        public void NewCase_EmptyActionRecipientFailsBeforeSettingsQuery(string recipient)
        {
            var parameters = new ParameterCollection
            {
                { "CaseDetails", "{}" },
                { "CaseID", "" },
                { "RecipientCompanyID", recipient }
            };
            var service = new SettingsService();
            var context = (IPluginExecutionContext)new PluginContextProxy(parameters).GetTransparentProxy();
            var ex = Assert.ThrowsException<InvalidPluginExecutionException>(() =>
                new PluginAsm::PostCasePlugin().Execute(new Services(service, context)));
            StringAssert.Contains(ex.Message, "Recipient company ID");
            Assert.IsNull(service.RecipientConditions);
        }

        [TestMethod]
        public void UnconfiguredRecipient_RemainsPlaintextEvenWhenAnotherIsConfigured()
        {
            var service = new SettingsService();
            service.Add("encryptedRecipient", "11", "https://example.test/keys");
            const string json = "{\"summary\":\"secret\"}";
            string result = E2eeCaseEncryption.PrepareForPost(new Services(service), PayloadType.Note, json, "ref", "22");
            Assert.AreEqual(json, result);
            CollectionAssert.AreEqual(new[] { "encryptedRecipient", "22" }, service.RecipientConditions);
        }

        [TestMethod]
        public void MatchingRecipientWithoutUri_FailsRatherThanPostingPlaintext()
        {
            var service = new SettingsService();
            service.Add("encryptedRecipient", "11", "https://example.test/other");
            service.Add("encryptedRecipient", "22", null);
            service.Add("Iss", "1", null);
            var ex = Assert.ThrowsException<InvalidPluginExecutionException>(() =>
                E2eeCaseEncryption.PrepareForPost(new Services(service), PayloadType.Note, "{}", "ref", "22"));
            StringAssert.Contains(ex.Message, "encryptedRecipient (ap_uri)");
            CollectionAssert.AreEqual(new[] { "encryptedRecipient", "22" }, service.RecipientConditions);
        }

        [TestMethod]
        public void MatchingRecipientWithUri_RequiresEncryptionConfiguration()
        {
            var service = new SettingsService();
            service.Add("encryptedRecipient", "11", null);
            service.Add("encryptedRecipient", "22", "https://example.test/keys");
            var ex = Assert.ThrowsException<InvalidPluginExecutionException>(() =>
                E2eeCaseEncryption.PrepareForPost(new Services(service), PayloadType.Note, "{}", "ref", "22"));
            StringAssert.Contains(ex.Message, "Iss setting (ap_value)");
            CollectionAssert.AreEqual(new[] { "encryptedRecipient", "22" }, service.RecipientConditions);
        }

        [TestMethod]
        public void MissingRecipientId_FailsBeforeQuery()
        {
            var service = new SettingsService();
            Assert.ThrowsException<InvalidPluginExecutionException>(() =>
                E2eeCaseEncryption.PrepareForPost(new Services(service), PayloadType.Note, "{}", "ref", null));
            Assert.IsNull(service.RecipientConditions);
        }

        private sealed class Services : IServiceProvider, IOrganizationServiceFactory
        {
            private readonly IOrganizationService service;
            private readonly IPluginExecutionContext context;
            public Services(IOrganizationService service, IPluginExecutionContext context = null)
            {
                this.service = service;
                this.context = context;
            }
            public object GetService(Type type)
            {
                if (type == typeof(IOrganizationServiceFactory)) return this;
                if (type == typeof(ITracingService)) return new TraceRecorder();
                if (type == typeof(IPluginExecutionContext)) return context;
                return null;
            }
            public IOrganizationService CreateOrganizationService(Guid? userId) { return service; }
        }

        private sealed class PluginContextProxy : RealProxy
        {
            private readonly ParameterCollection parameters;
            public PluginContextProxy(ParameterCollection parameters) : base(typeof(IPluginExecutionContext))
            {
                this.parameters = parameters;
            }
            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                object result;
                if (call.MethodName == "get_InputParameters") result = parameters;
                else if (call.MethodName == "get_UserId") result = Guid.Empty;
                else throw new NotSupportedException(call.MethodName);
                return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
            }
        }

        private sealed class SettingsService : IOrganizationService
        {
            private readonly List<Entity> settings = new List<Entity>();
            public string[] RecipientConditions;

            public void Add(string name, string value, string uri)
            {
                var setting = new Entity("ap_tsanetsetting");
                setting["ap_name"] = name;
                setting["ap_value"] = value;
                if (uri != null) setting["ap_uri"] = uri;
                settings.Add(setting);
            }

            public EntityCollection RetrieveMultiple(QueryBase queryBase)
            {
                var query = (QueryExpression)queryBase;
                if (query.EntityName == "systemuser")
                    return new EntityCollection(new List<Entity> { new Entity("systemuser", Guid.NewGuid()) });
                var name = query.Criteria.Conditions.Single(c => c.AttributeName == "ap_name").Values[0].ToString();
                var valueCondition = query.Criteria.Conditions.SingleOrDefault(c => c.AttributeName == "ap_value");
                string value = valueCondition == null ? null : valueCondition.Values[0].ToString();
                if (name == "encryptedRecipient") RecipientConditions = new[] { name, value };
                return new EntityCollection(settings.Where(s => s.GetAttributeValue<string>("ap_name") == name &&
                    (value == null || s.GetAttributeValue<string>("ap_value") == value)).Take(1).ToList());
            }
            public Guid Create(Entity entity) { throw new NotSupportedException(); }
            public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) { throw new NotSupportedException(); }
            public void Update(Entity entity) { throw new NotSupportedException(); }
            public void Delete(string entityName, Guid id) { throw new NotSupportedException(); }
            public OrganizationResponse Execute(OrganizationRequest request) { throw new NotSupportedException(); }
            public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException(); }
            public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException(); }
        }
    }
}
