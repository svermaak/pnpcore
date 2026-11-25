using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Test.Utilities;
using System;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Test.Base
{
    [TestClass]
    public class ListItemUserFieldMappingTests
    {
        [ClassInitialize]
        public static void TestFixtureSetup(TestContext context)
        {
            // All these tests are offline by design - no real SharePoint calls are made.
        }

        [TestMethod]
        public async Task HiddenUserFieldIsMappedFromStringIdCompanion_Issue1709()
        {
            using (var context = await TestCommon.Instance.GetContextWithoutInitializationAsync(TestCommon.TestSite))
            {
                var listItem = CreateListItem(context);

                // SharePoint does not return the "<field>Id" companion for a hidden user field, it only
                // returns "<field>StringId": a string for User, an array of strings for UserMulti.
                await ProcessAsync(listItem,
                    "{\"Id\":1,\"Title\":\"Item 1\"," +
                    "\"SingleUserHiddenStringId\":\"7\"," +
                    "\"MultiUserHiddenStringId\":[\"7\",\"8\"]}");

                // Before the fix neither field was mapped: the single user value ended up under the raw
                // "SingleUserHiddenStringId" key and the multi user value was dropped entirely.
                Assert.AreEqual(7, GetUserValue(listItem, "SingleUserHidden").LookupId);
                CollectionAssert.AreEqual(new[] { 7, 8 }, GetUserValues(listItem, "MultiUserHidden"));
            }
        }

        [TestMethod]
        public async Task VisibleUserFieldStillUsesIdCompanion_Issue1709()
        {
            using (var context = await TestCommon.Instance.GetContextWithoutInitializationAsync(TestCommon.TestSite))
            {
                var listItem = CreateListItem(context);

                // For a visible user field both companions are returned. The "StringId" values differ from
                // the "Id" values here only to show which companion the field value comes from.
                await ProcessAsync(listItem,
                    "{\"Id\":1,\"Title\":\"Item 1\"," +
                    "\"SingleUserVisibleId\":6,\"SingleUserVisibleStringId\":\"99\"," +
                    "\"MultiUserVisibleId\":[6],\"MultiUserVisibleStringId\":[\"99\"]}");

                Assert.AreEqual(6, GetUserValue(listItem, "SingleUserVisible").LookupId);
                CollectionAssert.AreEqual(new[] { 6 }, GetUserValues(listItem, "MultiUserVisible"));
            }
        }

        [TestMethod]
        public async Task LookupFieldStringIdCompanionIsNotMapped_Issue1709()
        {
            using (var context = await TestCommon.Instance.GetContextWithoutInitializationAsync(TestCommon.TestSite))
            {
                var listItem = CreateListItem(context);

                // FieldLookupValue does not read a string id, so a lookup must not be mapped from the
                // "StringId" companion: doing so would silently turn its value into -1.
                await ProcessAsync(listItem,
                    "{\"Id\":1,\"Title\":\"Item 1\"," +
                    "\"LookupHiddenStringId\":\"7\"}");

                Assert.IsFalse(listItem.Values.ContainsKey("LookupHidden"));
            }
        }

        private static ListItem CreateListItem(PnPContext context)
        {
            var list = new List();
            ((IDataModelWithContext)list).PnPContext = context;

            var fields = (FieldCollection)list.Fields;
            AddField(context, fields, "SingleUserVisible", FieldType.User, "User");
            AddField(context, fields, "SingleUserHidden", FieldType.User, "User");
            AddField(context, fields, "MultiUserVisible", FieldType.User, "UserMulti");
            AddField(context, fields, "MultiUserHidden", FieldType.User, "UserMulti");
            AddField(context, fields, "LookupHidden", FieldType.Lookup, "Lookup");

            var items = (ListItemCollection)list.Items;
            var listItem = new ListItem();
            ((IDataModelWithContext)listItem).PnPContext = context;
            ((IDataModelParent)listItem).Parent = items;
            items.Add(listItem);

            return listItem;
        }

        private static void AddField(PnPContext context, FieldCollection fields, string internalName, FieldType fieldTypeKind, string typeAsString)
        {
            var field = new Field
            {
                Id = Guid.NewGuid(),
                Title = internalName,
                InternalName = internalName,
                FieldTypeKind = fieldTypeKind,
                TypeAsString = typeAsString
            };
            ((IDataModelWithContext)field).PnPContext = context;
            fields.Add(field);
        }

        private static async Task ProcessAsync(ListItem listItem, string json)
        {
            using var jsonDoc = JsonDocument.Parse(json);
            var apiCall = new ApiCall("_api/dummy", ApiType.SPORest);
            var apiResponse = new ApiResponse(apiCall, jsonDoc.RootElement, Guid.NewGuid());

            await ((IDataModelProcess)listItem).ProcessResponseAsync(apiResponse);
        }

        private static IFieldUserValue GetUserValue(ListItem listItem, string fieldName)
        {
            Assert.IsTrue(listItem.Values.ContainsKey(fieldName), $"{fieldName} was not mapped");
            return (IFieldUserValue)listItem.Values[fieldName];
        }

        private static int[] GetUserValues(ListItem listItem, string fieldName)
        {
            Assert.IsTrue(listItem.Values.ContainsKey(fieldName), $"{fieldName} was not mapped");
            var collection = listItem.Values[fieldName] as IFieldValueCollection;
            Assert.IsNotNull(collection, $"{fieldName} was not mapped as a field value collection");
            return collection.Values.Cast<IFieldUserValue>().Select(p => p.LookupId).ToArray();
        }
    }
}
