using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live.Spikes
{
    /// <summary>
    /// Spike <b>S1</b>'s remaining tenant probe: can the Graph term store be made to preserve a
    /// caller-supplied term id?
    /// </summary>
    /// <remarks>
    /// <para>The static half of S1 is settled and needed no tenant: <c>ITermGroup.Id</c>,
    /// <c>ITermSet.Id</c> and <c>ITerm.Id</c> are all <c>{ get; }</c>, and no <c>AddAsync</c>
    /// overload accepts an id. PnP Core <em>cannot express</em> the operation. That is what
    /// triggered decision <b>D7</b> and the eight taxonomy CSOM requests.</para>
    /// <para><b>What is still open is one level down:</b> does the underlying Graph endpoint reject
    /// a supplied id, or does PnP Core simply not surface a parameter Graph would accept? Worth an
    /// hour, because if Graph accepts it the fallback can be dropped and eight requests deleted.</para>
    /// <para><b>Why this test is shaped the way it is.</b> A first version POSTed to
    /// <c>termStore/groups</c> and "passed" on a 400 — but the 400 was
    /// <c>Resource not found for the segment 'termStore'</c>, i.e. a wrong URL. It proved nothing
    /// while looking like evidence. The term store is <b>site-scoped</b>:
    /// <c>sites/{hostname},{siteId},{webId}/termstore/groups</c>.</para>
    /// <para>So every probe now <b>establishes that it reached the endpoint first</b>, with a GET
    /// that must succeed, and only then interprets what the POST did. A conclusion about Graph's
    /// behaviour is only drawn once the request demonstrably got there.</para>
    /// </remarks>
    [TestClass]
    public class S1GraphTermIdProbeTests : LiveTestBase
    {
        /// <summary>
        /// The id we try to impose. Fixed rather than random so a leaked object is recognisable.
        /// </summary>
        private static readonly Guid DesiredId = new Guid("5115f7d0-5115-4f7d-9115-f7d05115f7d0");

        /// <summary>
        /// Builds the site-scoped term store base path Graph actually uses.
        /// </summary>
        private static async Task<string> TermStoreBaseAsync(PnPContext context)
        {
            await context.Site.EnsurePropertiesAsync(s => s.Id).ConfigureAwait(false);
            await context.Web.EnsurePropertiesAsync(w => w.Id).ConfigureAwait(false);

            return $"sites/{context.Uri.DnsSafeHost},{context.Site.Id},{context.Web.Id}/termstore";
        }

        /// <summary>
        /// Proves the probe can reach the term store before anything is concluded from a failure.
        /// </summary>
        private static async Task AssertEndpointIsReachableAsync(PnPContext context, string termStoreBase)
        {
            try
            {
                ApiRequestResponse response = await context.Web.ExecuteRequestAsync(
                    new ApiRequest(ApiRequestType.Graph, $"{termStoreBase}/groups")).ConfigureAwait(false);

                Assert.IsFalse(string.IsNullOrEmpty(response.Response),
                    "The term store GET returned nothing - the probe cannot conclude anything about ids.");

                Console.WriteLine($"S1 probe: reached {termStoreBase}/groups OK - a POST failure from here is meaningful.");
            }
            catch (Exception ex)
            {
                Assert.Inconclusive(
                    $"Could not read the term store, so nothing can be concluded about supplied ids. " +
                    $"Most likely the account lacks term store admin rights.{Environment.NewLine}" +
                    $"Path: {termStoreBase}/groups{Environment.NewLine}{ex.Message}");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("S1")]
        public async Task Probe_DoesRawGraphAcceptACallerSuppliedTermGroupId()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                string termStoreBase = await TermStoreBaseAsync(context).ConfigureAwait(false);
                await AssertEndpointIsReachableAsync(context, termStoreBase).ConfigureAwait(false);

                await CleanUpLeakedTermGroupsAsync(context).ConfigureAwait(false);

                string groupName = $"{TestPrefix}S1Probe_{DateTime.UtcNow:yyyyMMddHHmmss}";
                string createdId = null;

                try
                {
                    string body = JsonSerializer.Serialize(new
                    {
                        id = DesiredId.ToString(),
                        displayName = groupName,
                        description = "S1 probe - safe to delete",
                    });

                    ApiRequestResponse response = await context.Web.ExecuteRequestAsync(
                        new ApiRequest(HttpMethod.Post, ApiRequestType.Graph, $"{termStoreBase}/groups", body))
                        .ConfigureAwait(false);

                    using (JsonDocument document = JsonDocument.Parse(response.Response))
                    {
                        createdId = document.RootElement.TryGetProperty("id", out JsonElement id) ? id.GetString() : null;
                    }

                    Console.WriteLine($"S1 probe - term GROUP: requested {DesiredId}, Graph ACCEPTED the call and returned {createdId ?? "<none>"}");

                    // Graph took the request. The question is now whether it honoured the id or
                    // quietly assigned its own - the latter being the more dangerous outcome, since
                    // it succeeds while silently breaking every template reference.
                    Assert.AreNotEqual(DesiredId.ToString(), createdId,
                        "Graph PRESERVED a caller-supplied term group id. That contradicts decision D7: the taxonomy " +
                        "CSOM fallback may be unnecessary and eight requests could be deleted. Re-open S1 before phase 6.");

                    Console.WriteLine("S1 probe - term GROUP: id was IGNORED and reassigned - confirms D7.");
                }
                catch (AssertFailedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    AssertRejectedForTheRightReason("term GROUP", ex);
                }
                finally
                {
                    await DeleteTermGroupDeepAsync(context, createdId).ConfigureAwait(false);
                    await CleanUpLeakedTermGroupsAsync(context).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("S1")]
        public async Task Probe_DoesRawGraphAcceptACallerSuppliedTermSetId()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                string termStoreBase = await TermStoreBaseAsync(context).ConfigureAwait(false);
                await AssertEndpointIsReachableAsync(context, termStoreBase).ConfigureAwait(false);

                string groupName = $"{TestPrefix}S1SetProbe_{DateTime.UtcNow:yyyyMMddHHmmss}";
                ITermGroup group = null;
                string createdSetId = null;

                try
                {
                    // A term set needs a group to live in; create that the normal way.
                    group = await context.TermStore.Groups.AddAsync(groupName, "S1 probe - safe to delete")
                        .ConfigureAwait(false);

                    string body = JsonSerializer.Serialize(new
                    {
                        id = DesiredId.ToString(),
                        localizedNames = new[] { new { languageTag = "en-US", name = $"{TestPrefix}Set" } },
                        parentGroup = new { id = group.Id },
                    });

                    ApiRequestResponse response = await context.Web.ExecuteRequestAsync(
                        new ApiRequest(HttpMethod.Post, ApiRequestType.Graph, $"{termStoreBase}/sets", body))
                        .ConfigureAwait(false);

                    using (JsonDocument document = JsonDocument.Parse(response.Response))
                    {
                        createdSetId = document.RootElement.TryGetProperty("id", out JsonElement id) ? id.GetString() : null;
                    }

                    Console.WriteLine($"S1 probe - term SET: requested {DesiredId}, Graph ACCEPTED the call and returned {createdSetId ?? "<none>"}");

                    Assert.AreNotEqual(DesiredId.ToString(), createdSetId,
                        "Graph PRESERVED a caller-supplied term set id. That contradicts decision D7 - re-open S1.");

                    Console.WriteLine("S1 probe - term SET: id was IGNORED and reassigned - confirms D7.");
                }
                catch (AssertFailedException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    AssertRejectedForTheRightReason("term SET", ex);
                }
                finally
                {
                    if (group != null)
                    {
                        try
                        {
                            // Deep delete: this probe puts a term set inside the group, and Graph
                            // refuses to delete a non-empty one.
                            await DeleteTermGroupDeepAsync(group).ConfigureAwait(false);
                        }
                        catch (Exception)
                        {
                            // Swept up by CleanUpLeakedTermGroupsAsync on the next run.
                        }
                    }
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("S1")]
        public async Task Probe_ConfirmsTheModelSurfaceCannotExpressASuppliedId()
        {
            // The static finding, re-asserted against the live assembly so it cannot rot: there is
            // no AddAsync overload taking an id, and Id is read-only. Needs no network, but belongs
            // with the probe it justifies.
            Assert.IsFalse(typeof(ITermGroup).GetProperty(nameof(ITermGroup.Id)).CanWrite,
                "ITermGroup.Id became settable - re-open S1.");
            Assert.IsFalse(typeof(ITermSet).GetProperty(nameof(ITermSet.Id)).CanWrite,
                "ITermSet.Id became settable - re-open S1.");
            Assert.IsFalse(typeof(ITerm).GetProperty(nameof(ITerm.Id)).CanWrite,
                "ITerm.Id became settable - re-open S1.");

            await Task.CompletedTask.ConfigureAwait(false);
        }

        /// <summary>
        /// Confirms Graph refused the supplied id for a reason that actually settles the question.
        /// </summary>
        /// <remarks>
        /// <para>"Graph threw" is not evidence on its own, and this method exists because the first
        /// version of this probe proved exactly that: it passed on a 400 that turned out to be a
        /// wrong URL.</para>
        /// <para>Reaching the endpoint is established separately by
        /// <see cref="AssertEndpointIsReachableAsync"/>. What is checked here is that the refusal
        /// was not an auth failure and not a routing error - i.e. Graph understood the request and
        /// refused the <em>payload</em>.</para>
        /// </remarks>
        private static void AssertRejectedForTheRightReason(string what, Exception ex)
        {
            var serviceException = ex as ServiceException;
            var error = serviceException?.Error as ServiceError;

            int httpCode = error?.HttpResponseCode ?? 0;
            string graphCode = error?.Code;
            string detail = error?.Message ?? ex.Message;

            Console.WriteLine($"S1 probe - {what}: rejected. HTTP {httpCode}, code '{graphCode}'{Environment.NewLine}  {detail}");

            Assert.AreNotEqual(401, httpCode,
                $"The probe was not authenticated, so it proved nothing about {what} ids.");

            Assert.AreNotEqual(403, httpCode,
                $"The probe was forbidden - most likely no term store admin rights - so it proved nothing about {what} ids.");

            // The trap that caught the first version: a 400 whose message is about routing means
            // the request never reached the behaviour under test.
            Assert.IsFalse(detail != null && detail.Contains("Resource not found for the segment", StringComparison.OrdinalIgnoreCase),
                $"The URL was wrong, so this proved nothing about {what} ids. Detail: {detail}");

            Assert.IsTrue(httpCode >= 400 && httpCode < 500,
                $"Expected a 4xx refusal of the payload, got HTTP {httpCode}.");

            Console.WriteLine($"S1 probe - {what}: Graph REFUSED the supplied id - confirms D7.");
        }

        private static async Task DeleteGroupIfPresentAsync(PnPContext context, string groupId)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                return;
            }

            try
            {
                ITermGroup group = await context.TermStore.Groups.GetByIdAsync(groupId).ConfigureAwait(false);
                if (group != null)
                {
                    await DeleteTermGroupDeepAsync(group).ConfigureAwait(false);
                }
            }
            catch (Exception)
            {
                // Best effort - the prefix sweep will catch it next run.
            }
        }
    }
}
