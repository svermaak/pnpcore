using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.Audit;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.InformationPolicy;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.Publishing;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources;
using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.Requests;
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace PnP.Core.Provisioning.Test.Offline.Csom
{
    /// <summary>
    /// Asserts the exact XML each CSOM request puts on the wire.
    /// </summary>
    /// <remarks>
    /// <para>A CSOM request cannot be type-checked. The server type ids are GUIDs, the method names
    /// are strings, and a mistake in either fails at runtime with an error that names neither -
    /// which makes these the highest-value offline tests in the project.</para>
    /// <para>They exercise the real serializer (<see cref="CSOMApiCallBuilder"/>), so what is
    /// asserted is genuinely what SharePoint would receive. What they cannot prove is that
    /// SharePoint <em>accepts</em> it; that is the live round-trip test each request still needs.</para>
    /// </remarks>
    [TestClass]
    public class CsomPayloadTests
    {
        private static readonly Guid SiteId = new Guid("11111111-1111-1111-1111-111111111111");
        private static readonly Guid WebId = new Guid("22222222-2222-2222-2222-222222222222");

        /// <summary>
        /// Serializes a request exactly as the SDK would before posting it.
        /// </summary>
        private static string Serialize<T>(IRequest<T> request)
        {
            var builder = new CSOMApiCallBuilder();
            builder.AddRequest(request);
            return builder.SerializeCSOMRequests();
        }

        /// <summary>
        /// Every payload must be well-formed XML with the CSOM envelope. A request that produces
        /// malformed XML fails here rather than as an opaque server error.
        /// </summary>
        private static XDocument ParseAndValidateEnvelope(string payload)
        {
            XDocument document = XDocument.Parse(payload);

            Assert.AreEqual("Request", document.Root.Name.LocalName);
            Assert.IsNotNull(document.Root.Element(document.Root.Name.Namespace + "Actions"),
                "Every CSOM request needs an <Actions> element.");
            Assert.IsNotNull(document.Root.Element(document.Root.Name.Namespace + "ObjectPaths"),
                "Every CSOM request needs an <ObjectPaths> element.");

            return document;
        }

        #region T21 - audit settings

        [TestMethod]
        [TestCategory("Offline")]
        public void GetAuditRequest_AddressesTheSiteAuditAndQueriesAuditFlags()
        {
            string payload = Serialize(new GetAuditRequest(SiteId, WebId));
            ParseAndValidateEnvelope(payload);

            // The audit hangs off the *site*, not the web - so the identity carries the trailing
            // :site: segment.
            StringAssert.Contains(payload, $"site:{SiteId}:web:{WebId}:site:{SiteId}");
            StringAssert.Contains(payload, "Name=\"Audit\"");
            StringAssert.Contains(payload, "Name=\"AuditFlags\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void UpdateAuditRequest_SetsFlagsAsAnIntegerAndCallsUpdate()
        {
            // The delete flag is called ObjectDelete, not Delete.
            string payload = Serialize(new UpdateAuditRequest(SiteId, WebId, AuditMaskType.CheckOut | AuditMaskType.ObjectDelete));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, "<SetProperty");
            StringAssert.Contains(payload, "Name=\"AuditFlags\"");

            // CheckOut(1) | ObjectDelete(8) == 9. The enum has no [Flags] attribute but is combined with
            // bitwise or anyway - reproduced from the real CSOM enum in phase 1 - so the wire value
            // must be the numeric combination, not a name list.
            StringAssert.Contains(payload, ">9<");
            StringAssert.Contains(payload, "<Method Name=\"Update\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void UpdateAuditRequest_SerializesAllAsMinusOne()
        {
            // AuditMaskType.All is -1, not the sum of its flags. Sending 15871 - the sum - would
            // silently enable a different set.
            string payload = Serialize(new UpdateAuditRequest(SiteId, WebId, AuditMaskType.All));

            StringAssert.Contains(payload, ">-1<");
        }

        #endregion

        #region T22 - site policy

        [TestMethod]
        [TestCategory("Offline")]
        public void GetProjectPoliciesRequest_UsesTheProjectPolicyTypeIdAndPassesTheWebByReference()
        {
            string payload = Serialize(new GetProjectPoliciesRequest(SiteId, WebId));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, CsomTypeIds.ProjectPolicy);
            StringAssert.Contains(payload, "Name=\"GetProjectPolicies\"");

            // The ClientRuntimeContext parameter is implicit and must NOT be serialized; only the
            // web is sent, as an object reference.
            StringAssert.Contains(payload, "<Parameter ObjectPathId=");
            StringAssert.Contains(payload, "_Child_Items_".Length > 0 ? "ChildItemQuery" : "");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetCurrentlyAppliedProjectPolicyRequest_UsesTheOnWebMethodName()
        {
            string payload = Serialize(new GetCurrentlyAppliedProjectPolicyRequest(SiteId, WebId));
            ParseAndValidateEnvelope(payload);

            // The CSOM method is GetCurrentlyAppliedProjectPolicyOnWeb - the "OnWeb" suffix is easy
            // to drop and the failure would be a runtime "method not found".
            StringAssert.Contains(payload, "Name=\"GetCurrentlyAppliedProjectPolicyOnWeb\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ApplyProjectPolicyRequest_ResolvesThePolicyByIndexThenAppliesIt()
        {
            string payload = Serialize(new ApplyProjectPolicyRequest(SiteId, WebId, 2));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, "Name=\"GetProjectPolicies\"");
            StringAssert.Contains(payload, "Name=\"GetItemAtIndex\"");
            StringAssert.Contains(payload, "Name=\"ApplyProjectPolicy\"");

            // A void static is a <StaticMethod> action, not an <ObjectPath> to a static method path.
            StringAssert.Contains(payload, "<StaticMethod Id=");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ApplyProjectPolicyRequest_RejectsANegativeIndex()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                new ApplyProjectPolicyRequest(SiteId, WebId, -1));
        }

        #endregion

        #region T20 - image renditions

        [TestMethod]
        [TestCategory("Offline")]
        public void GetImageRenditionsRequest_SendsNoParameters()
        {
            string payload = Serialize(new GetImageRenditionsRequest());
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, CsomTypeIds.SiteImageRenditions);
            StringAssert.Contains(payload, "Name=\"GetRenditions\"");

            // GetRenditions(context) takes only the implicit context, so the parameter list is empty.
            StringAssert.Contains(payload, "<Parameters></Parameters>");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SetImageRenditionsRequest_ConstructsEachRenditionAndPassesThemAsAnObjectArray()
        {
            var renditions = new List<ImageRenditionInfo>
            {
                new ImageRenditionInfo { Name = "Banner", Width = 940, Height = 300 },
                new ImageRenditionInfo { Name = "Thumbnail", Width = 100, Height = 100 },
            };

            string payload = Serialize(new SetImageRenditionsRequest(renditions));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, CsomTypeIds.ImageRendition);
            StringAssert.Contains(payload, "Name=\"SetRenditions\"");
            StringAssert.Contains(payload, "Banner");
            StringAssert.Contains(payload, "Thumbnail");
            StringAssert.Contains(payload, ">940<");

            // Each rendition is a client object, so the array carries object references rather
            // than values.
            StringAssert.Contains(payload, "<Parameter Type=\"Array\"><Object ObjectPathId=");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SetImageRenditionsRequest_RejectsNull()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new SetImageRenditionsRequest(null));
        }

        #endregion

        #region T6 - user resources

        [TestMethod]
        [TestCategory("Offline")]
        public void SetValueForUICultureRequest_WalksFromTheParentIdentityToTheResourceProperty()
        {
            var resource = UserResourcePath.ForWeb(SiteId, WebId, ResourceProperty.Title);
            string payload = Serialize(new SetValueForUICultureRequest(resource, "nl-NL", "Mijn site"));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, $"site:{SiteId}:web:{WebId}");
            StringAssert.Contains(payload, "Name=\"TitleResource\"");
            StringAssert.Contains(payload, "Name=\"SetValueForUICulture\"");
            StringAssert.Contains(payload, "nl-NL");
            StringAssert.Contains(payload, "Mijn site");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetValueForUICultureRequest_PassesOnlyTheCultureName()
        {
            var resource = UserResourcePath.ForWeb(SiteId, WebId, ResourceProperty.Description);
            string payload = Serialize(new GetValueForUICultureRequest(resource, "de-DE"));
            ParseAndValidateEnvelope(payload);

            StringAssert.Contains(payload, "Name=\"DescriptionResource\"");
            StringAssert.Contains(payload, "Name=\"GetValueForUICulture\"");
            StringAssert.Contains(payload, "de-DE");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void UserResourceRequests_RejectAMissingCultureName()
        {
            var resource = UserResourcePath.ForWeb(SiteId, WebId, ResourceProperty.Title);

            Assert.ThrowsException<ArgumentException>(() => new SetValueForUICultureRequest(resource, null, "x"));
            Assert.ThrowsException<ArgumentException>(() => new GetValueForUICultureRequest(resource, ""));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SetValueForUICultureRequest_AlwaysUpdatesTheParent()
        {
            // Guards a bug found only by a live test: SetValueForUICulture merely STAGES the
            // change. Without an Update on the PARENT it is silently discarded - the call
            // succeeds, nothing is raised, and a read-back returns the object's default title.
            // Localization would have failed silently across a dozen phase 6 handlers.
            //
            // The payload test could not have caught the original defect (the XML was structurally
            // valid) but it can stop the fix being refactored away.
            var web = UserResourcePath.ForWeb(SiteId, WebId, ResourceProperty.Title);
            string webPayload = Serialize(new SetValueForUICultureRequest(web, "de-DE", "Titel"));

            StringAssert.Contains(webPayload, "Name=\"SetValueForUICulture\"");
            StringAssert.Contains(webPayload, "<Method Name=\"Update\"",
                "A web resource change must be followed by Web.Update() or it is discarded.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SetValueForUICultureRequest_UsesTheRightUpdateMethodPerParentType()
        {
            // The persist call differs by parent, and getting it wrong fails silently:
            //   Web, List   -> Update()
            //   ContentType -> Update(true)                 (pushes to child content types)
            //   Field       -> UpdateAndPushChanges(true)   (pushes to list columns)
            string list = Serialize(new SetValueForUICultureRequest(
                UserResourcePath.ForList(SiteId, WebId, Guid.NewGuid(), ResourceProperty.Title), "de-DE", "x"));
            StringAssert.Contains(list, "<Method Name=\"Update\"");

            string field = Serialize(new SetValueForUICultureRequest(
                UserResourcePath.ForField(SiteId, WebId, Guid.NewGuid(), ResourceProperty.Title), "de-DE", "x"));
            StringAssert.Contains(field, "<Method Name=\"UpdateAndPushChanges\"",
                "A site column needs UpdateAndPushChanges so list columns derived from it follow.");
            StringAssert.Contains(field, "Type=\"Boolean\">true<");

            string contentType = Serialize(new SetValueForUICultureRequest(
                UserResourcePath.ForContentType(SiteId, WebId, "0x0101", ResourceProperty.ContentTypeName), "de-DE", "x"));
            StringAssert.Contains(contentType, "<Method Name=\"Update\"");
            StringAssert.Contains(contentType, "Type=\"Boolean\">true<",
                "A content type needs Update(true) so child content types follow.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetValueForUICultureRequest_DoesNotUpdateTheParent()
        {
            // A read must not write. An Update here would dirty the object on every extract.
            string payload = Serialize(new GetValueForUICultureRequest(
                UserResourcePath.ForWeb(SiteId, WebId, ResourceProperty.Title), "de-DE"));

            StringAssert.Contains(payload, "Name=\"GetValueForUICulture\"");
            Assert.IsFalse(payload.Contains("<Method Name=\"Update\""),
                "Reading a localized value must not update the parent.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ContentTypeDisplayNameUsesNameResourceNotTitleResource()
        {
            // The content type property is NameResource. Using TitleResource silently fails to
            // localize content types, which is the sort of thing that only shows up in a
            // multilingual tenant weeks later.
            Assert.AreEqual("NameResource", ResourceProperty.ContentTypeName);
            Assert.AreEqual("TitleResource", ResourceProperty.Title);
        }

        #endregion

        #region Identity strings

        [TestMethod]
        [TestCategory("Offline")]
        public void CsomIdentity_ProducesTheDocumentedShapes()
        {
            var listId = new Guid("33333333-3333-3333-3333-333333333333");

            StringAssert.EndsWith(CsomIdentity.Web(SiteId, WebId), $":site:{SiteId}:web:{WebId}");

            // A site identity still routes through the web that reached it, then names the site again.
            StringAssert.EndsWith(CsomIdentity.Site(SiteId, WebId), $":site:{SiteId}:web:{WebId}:site:{SiteId}");

            StringAssert.EndsWith(CsomIdentity.List(SiteId, WebId, listId), $":web:{WebId}:list:{listId}");
            StringAssert.EndsWith(CsomIdentity.ContentType(SiteId, WebId, "0x0101"), ":contenttype:0x0101");

            // Every identity carries the fixed client object model application id.
            StringAssert.Contains(CsomIdentity.Web(SiteId, WebId), "740c6a0b-85e2-48a0-a494-e0f1759d4aa7");
        }

        #endregion

        #region The reference-type constraint

        [TestMethod]
        [TestCategory("Offline")]
        public void EveryCsomRequestResultTypeIsAReferenceType()
        {
            // PnP Core batches requests by casting each to IRequest<object>. IRequest<out T> is
            // covariant, but covariance does not apply to value types - so a request declared as
            // IRequest<int> or IRequest<SomeEnum?> compiles cleanly and then throws
            // InvalidCastException at run time, from inside the SDK, with nothing in the stack
            // naming the offending request.
            //
            // None of PnP Core's own 17 requests returns a value type, so the constraint is
            // undocumented there and there is no compile-time check. This is that check.
            var offenders = new List<string>();

            foreach (Type type in typeof(CsomTypeIds).Assembly.GetTypes())
            {
                if (type.IsAbstract || type.IsInterface) continue;

                foreach (Type iface in type.GetInterfaces())
                {
                    if (!iface.IsGenericType || iface.GetGenericTypeDefinition() != typeof(IRequest<>)) continue;

                    Type resultType = iface.GetGenericArguments()[0];
                    if (resultType.IsValueType)
                    {
                        offenders.Add($"{type.Name} -> IRequest<{resultType.Name}>");
                    }
                }
            }

            Assert.AreEqual(0, offenders.Count,
                "These requests declare a value-type result and will throw InvalidCastException when sent. " +
                "Wrap the value in a small class, as AuditSettingsInfo does:" +
                Environment.NewLine + string.Join(Environment.NewLine, offenders));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetAuditRequest_ReturnsAReferenceTypedResult()
        {
            // Guards the specific fix: this request originally declared IRequest<AuditMaskType?>.
            Assert.IsFalse(typeof(AuditSettingsInfo).IsValueType);
        }

        #endregion

        #region Type ids

        [TestMethod]
        [TestCategory("Offline")]
        public void TaxonomySessionTypeIdMatchesThePnPCoreValue()
        {
            // PnP Core's own GetParentOfTermRequest hardcodes this GUID. That it matches what was
            // extracted from the CSOM assembly is what validates the extraction method for all the
            // other type ids, none of which have a second source to check against.
            Assert.AreEqual("{981cbc68-9edc-4f8d-872f-71146fcbb84f}", CsomTypeIds.TaxonomySession);
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void EveryTypeIdIsAWellFormedBracedGuid()
        {
            foreach (System.Reflection.FieldInfo field in typeof(CsomTypeIds)
                .GetFields(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static))
            {
                if (field.FieldType != typeof(string)) continue;

                string value = (string)field.GetRawConstantValue();

                Assert.IsTrue(value.StartsWith("{", StringComparison.Ordinal) && value.EndsWith("}", StringComparison.Ordinal),
                    $"{field.Name} must be brace-wrapped - CSOM rejects a bare GUID.");
                Assert.IsTrue(Guid.TryParse(value, out _), $"{field.Name} is not a valid GUID: {value}");
            }
        }

        #endregion
    }
}
