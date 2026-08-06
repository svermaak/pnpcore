using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.WebParts;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.Workflows;
using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.Requests;
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace PnP.Core.Provisioning.Test.Offline.Csom
{
    /// <summary>
    /// Payload tests for the classic web part (<b>T18</b>) and SharePoint 2013 workflow
    /// (<b>T16</b>) CSOM requests.
    /// </summary>
    [TestClass]
    public class CsomWebPartAndWorkflowPayloadTests
    {
        private static readonly Guid SiteId = new Guid("11111111-1111-1111-1111-111111111111");
        private static readonly Guid WebId = new Guid("22222222-2222-2222-2222-222222222222");
        private static readonly Guid WebPartId = new Guid("dddddddd-dddd-dddd-dddd-dddddddddddd");
        private static readonly Guid DefinitionId = new Guid("eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee");
        private const string PageUrl = "/sites/team/SitePages/Home.aspx";

        private static string Serialize<T>(IRequest<T> request)
        {
            var builder = new CSOMApiCallBuilder();
            builder.AddRequest(request);
            string payload = builder.SerializeCSOMRequests();
            XDocument.Parse(payload);
            return payload;
        }

        #region T18 - web parts

        [TestMethod]
        [TestCategory("Offline")]
        public void AddWebPartRequest_ImportsThenAddsInOneBatch()
        {
            string payload = Serialize(new AddWebPartRequest(SiteId, WebId, PageUrl,
                "<webParts>...</webParts>", "Left", 0));

            StringAssert.Contains(payload, "Name=\"GetFileByServerRelativeUrl\"");
            StringAssert.Contains(payload, PageUrl);
            StringAssert.Contains(payload, "Name=\"GetLimitedWebPartManager\"");

            // Import and add must share a batch: ImportWebPart returns a server-side object that
            // cannot be carried across round trips.
            StringAssert.Contains(payload, "Name=\"ImportWebPart\"");
            StringAssert.Contains(payload, "Name=\"AddWebPart\"");

            // AddWebPart takes the WebPart, not the WebPartDefinition that ImportWebPart returned.
            StringAssert.Contains(payload, "Name=\"WebPart\"");
            StringAssert.Contains(payload, "Left");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void WebPartRequests_UseTheSharedPersonalizationScope()
        {
            // Personal scope would make the customization visible only to the account running the
            // template - silently useless.
            string payload = Serialize(new GetWebPartDefinitionsRequest(SiteId, WebId, PageUrl));

            StringAssert.Contains(payload, "Name=\"GetLimitedWebPartManager\"");
            StringAssert.Contains(payload, "Type=\"Enum\">1<");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetWebPartDefinitionsRequest_QueriesTheWebPartsCollection()
        {
            string payload = Serialize(new GetWebPartDefinitionsRequest(SiteId, WebId, PageUrl));

            StringAssert.Contains(payload, "Name=\"WebParts\"");
            StringAssert.Contains(payload, "ChildItemQuery");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void MoveWebPartToRequest_ResolvesTheWebPartByIdThenMovesIt()
        {
            string payload = Serialize(new MoveWebPartToRequest(SiteId, WebId, PageUrl, WebPartId, "Right", 2));

            StringAssert.Contains(payload, "Name=\"GetById\"");
            StringAssert.Contains(payload, WebPartId.ToString());
            StringAssert.Contains(payload, "Name=\"MoveWebPartTo\"");
            StringAssert.Contains(payload, "Right");
            StringAssert.Contains(payload, ">2<");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void DeleteWebPartRequest_ResolvesTheWebPartByIdThenDeletesIt()
        {
            string payload = Serialize(new DeleteWebPartRequest(SiteId, WebId, PageUrl, WebPartId));

            StringAssert.Contains(payload, "Name=\"GetById\"");
            StringAssert.Contains(payload, "Name=\"DeleteWebPart\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SaveWebPartPropertiesRequest_SetsPropertiesThenSaves()
        {
            string payload = Serialize(new SaveWebPartPropertiesRequest(SiteId, WebId, PageUrl, WebPartId,
                title: "My Web Part", zoneIndex: 3));

            StringAssert.Contains(payload, "Name=\"Title\"");
            StringAssert.Contains(payload, "My Web Part");
            StringAssert.Contains(payload, "Name=\"ZoneIndex\"");

            // The setters stage the change; SaveWebPartChanges is what persists it. Same trap as
            // Audit.Update() and the term store's CommitAll().
            StringAssert.Contains(payload, "Name=\"SaveWebPartChanges\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void WebPartRequests_RejectAMissingFileUrl()
        {
            Assert.ThrowsException<ArgumentException>(() => new AddWebPartRequest(SiteId, WebId, null, "<x/>", "Left", 0));
            Assert.ThrowsException<ArgumentException>(() => new GetWebPartDefinitionsRequest(SiteId, WebId, ""));
            Assert.ThrowsException<ArgumentException>(() => new DeleteWebPartRequest(SiteId, WebId, null, WebPartId));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void AddWebPartRequest_RejectsMissingXml()
        {
            Assert.ThrowsException<ArgumentException>(() => new AddWebPartRequest(SiteId, WebId, PageUrl, null, "Left", 0));
        }

        #endregion

        #region T16 - workflows

        [TestMethod]
        [TestCategory("Offline")]
        public void WorkflowRequests_ConstructTheServicesManagerFromTheWeb()
        {
            string payload = Serialize(new GetWorkflowDefinitionsRequest(SiteId, WebId));

            StringAssert.Contains(payload, CsomTypeIds.WorkflowServicesManager);
            StringAssert.Contains(payload, $"site:{SiteId}:web:{WebId}");
            StringAssert.Contains(payload, "<Constructor");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetWorkflowDefinitionsRequest_EnumeratesThroughTheDeploymentService()
        {
            string payload = Serialize(new GetWorkflowDefinitionsRequest(SiteId, WebId, publishedOnly: true));

            StringAssert.Contains(payload, "Name=\"GetWorkflowDeploymentService\"");
            StringAssert.Contains(payload, "Name=\"EnumerateDefinitions\"");
            StringAssert.Contains(payload, ">true<");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SaveWorkflowDefinitionRequest_SendsTheXamlAndSaves()
        {
            var definition = new WorkflowDefinitionInfo
            {
                DisplayName = "Approval",
                Description = "An approval workflow",
                Xaml = "<Activity>...</Activity>",
            };

            string payload = Serialize(new SaveWorkflowDefinitionRequest(SiteId, WebId, definition));

            StringAssert.Contains(payload, CsomTypeIds.WorkflowDefinition);
            StringAssert.Contains(payload, "Name=\"Xaml\"");
            StringAssert.Contains(payload, "Name=\"SaveDefinition\"");
            StringAssert.Contains(payload, "Approval");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SaveWorkflowDefinitionRequest_SendsTheIdOnlyWhenUpdating()
        {
            // Supplying an id updates an existing definition; omitting it creates a new one. A
            // template re-applied to the same site must update, or every run leaves another copy.
            var withId = new WorkflowDefinitionInfo { Id = DefinitionId, Xaml = "<x/>" };
            StringAssert.Contains(Serialize(new SaveWorkflowDefinitionRequest(SiteId, WebId, withId)),
                DefinitionId.ToString());

            var withoutId = new WorkflowDefinitionInfo { Xaml = "<x/>" };
            Assert.IsFalse(Serialize(new SaveWorkflowDefinitionRequest(SiteId, WebId, withoutId)).Contains("Name=\"Id\""),
                "An empty id must not be sent - it would be taken as a real id.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void SaveWorkflowDefinitionRequest_RequiresXaml()
        {
            Assert.ThrowsException<ArgumentException>(() =>
                new SaveWorkflowDefinitionRequest(SiteId, WebId, new WorkflowDefinitionInfo { DisplayName = "x" }));

            Assert.ThrowsException<ArgumentNullException>(() =>
                new SaveWorkflowDefinitionRequest(SiteId, WebId, null));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void PublishAndDeleteDefinitionRequests_TargetTheDeploymentService()
        {
            StringAssert.Contains(Serialize(new PublishWorkflowDefinitionRequest(SiteId, WebId, DefinitionId)),
                "Name=\"PublishDefinition\"");

            StringAssert.Contains(Serialize(new DeleteWorkflowDefinitionRequest(SiteId, WebId, DefinitionId)),
                "Name=\"DeleteDefinition\"");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetWorkflowSubscriptionsRequest_NarrowsToAListWhenAsked()
        {
            var listId = new Guid("ffffffff-ffff-ffff-ffff-ffffffffffff");

            StringAssert.Contains(Serialize(new GetWorkflowSubscriptionsRequest(SiteId, WebId)),
                "Name=\"EnumerateSubscriptions\"");

            string byList = Serialize(new GetWorkflowSubscriptionsRequest(SiteId, WebId, listId));
            StringAssert.Contains(byList, "Name=\"EnumerateSubscriptionsByList\"");
            StringAssert.Contains(byList, listId.ToString());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void PublishWorkflowSubscriptionRequest_PicksTheListOrSiteVariant()
        {
            var subscription = new WorkflowSubscriptionInfo
            {
                DefinitionId = DefinitionId,
                Name = "On item added",
                Enabled = true,
            };

            string site = Serialize(new PublishWorkflowSubscriptionRequest(SiteId, WebId, subscription));
            StringAssert.Contains(site, "Name=\"PublishSubscription\"");

            var listId = new Guid("ffffffff-ffff-ffff-ffff-ffffffffffff");
            string list = Serialize(new PublishWorkflowSubscriptionRequest(SiteId, WebId, subscription, listId));
            StringAssert.Contains(list, "Name=\"PublishSubscriptionForList\"");
            StringAssert.Contains(list, listId.ToString());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void PublishWorkflowSubscriptionRequest_SendsAssociationProperties()
        {
            var subscription = new WorkflowSubscriptionInfo
            {
                DefinitionId = DefinitionId,
                PropertyDefinitions = new Dictionary<string, string>
                {
                    { "TaskListId", "{tasklist}" },
                    { "HistoryListId", "{historylist}" },
                },
            };

            string payload = Serialize(new PublishWorkflowSubscriptionRequest(SiteId, WebId, subscription));

            // PropertyDefinitions is a dictionary, which CSOM cannot assign wholesale - each entry
            // is its own SetProperty call.
            StringAssert.Contains(payload, "Name=\"SetProperty\"");
            StringAssert.Contains(payload, "TaskListId");
            StringAssert.Contains(payload, "HistoryListId");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void PublishWorkflowSubscriptionRequest_RequiresADefinitionId()
        {
            Assert.ThrowsException<ArgumentException>(() =>
                new PublishWorkflowSubscriptionRequest(SiteId, WebId, new WorkflowSubscriptionInfo()));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void EnumerateWorkflowInstancesRequest_CountsThroughTheInstanceService()
        {
            string payload = Serialize(new EnumerateWorkflowInstancesRequest(SiteId, WebId, DefinitionId));

            StringAssert.Contains(payload, "Name=\"GetWorkflowInstanceService\"");
            StringAssert.Contains(payload, "Name=\"GetSubscription\"");
            StringAssert.Contains(payload, "Name=\"CountInstances\"");
        }

        #endregion
    }
}
