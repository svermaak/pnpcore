using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using PnP.Core.Services.Core.CSOM;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Web
{
    /// <summary>
    /// Writes the web's regional settings.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T1</b>. This exists because <b>REST cannot do it</b>: a <c>MERGE</c> against
    /// <c>_api/web/regionalsettings</c> is refused with <i>"The type SP.RegionalSettings does not
    /// support HTTP PATCH method"</i>. <c>SP.RegionalSettings</c> is not an updatable REST entity,
    /// so CSOM is the only route - which is what PnP Framework always used.</para>
    /// <para><b>The commit is on the settings object, not the web.</b> <c>SP.RegionalSettings</c>
    /// carries its own <c>Update()</c>. Calling <c>Web.Update()</c> instead would leave the change
    /// staged and silently discarded - the same trap as <c>Audit.Update()</c>, the term store's
    /// <c>CommitAll()</c> and the user resource parent update.</para>
    /// <para><b>Only supplied properties are written.</b> A template that mentions three settings
    /// must not reset the other nine to zero, so each is emitted only when present in the
    /// dictionary.</para>
    /// <para>Types matter: CSOM declares most of these as <c>Int16</c>, but <c>FirstDayOfWeek</c>
    /// and <c>LocaleId</c> as <c>UInt32</c>, and <c>ShowWeeks</c>/<c>Time24</c> as <c>Boolean</c>.
    /// Sending the wrong CSOM type name fails the whole request.
    /// </para>
    /// </remarks>
    internal sealed class SetRegionalSettingsRequest : IRequest<object>
    {
        /// <summary>
        /// The CSOM type name each settable property expects.
        /// </summary>
        /// <remarks>
        /// Taken from the shipped CSOM assembly rather than guessed - <c>Int16</c> versus
        /// <c>UInt32</c> is not inferable from the property's meaning.
        /// </remarks>
        private static readonly Dictionary<string, string> PropertyTypes = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["AdjustHijriDays"] = "Int16",
            ["AlternateCalendarType"] = "Int16",
            ["CalendarType"] = "Int16",
            ["Collation"] = "Int16",
            ["FirstDayOfWeek"] = "UInt32",
            ["FirstWeekOfYear"] = "Int16",
            ["LocaleId"] = "UInt32",
            ["ShowWeeks"] = "Boolean",
            ["Time24"] = "Boolean",
            ["WorkDayEndHour"] = "Int16",
            ["WorkDayStartHour"] = "Int16",
            ["WorkDays"] = "Int16",
        };

        private readonly Guid siteId;
        private readonly Guid webId;
        private readonly IDictionary<string, object> properties;

        internal SetRegionalSettingsRequest(Guid siteId, Guid webId, IDictionary<string, object> properties)
        {
            if (properties == null || properties.Count == 0)
            {
                throw new ArgumentException("At least one regional setting is required.", nameof(properties));
            }

            foreach (string name in properties.Keys)
            {
                if (!PropertyTypes.ContainsKey(name))
                {
                    throw new ArgumentException(
                        $"'{name}' is not a settable regional setting. Settable: {string.Join(", ", PropertyTypes.Keys)}.",
                        nameof(properties));
                }
            }

            this.siteId = siteId;
            this.webId = webId;
            this.properties = properties;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int webIdentityId = idProvider.GetActionId();
            int regionalSettingsId = idProvider.GetActionId();

            // Reach the settings through the web's identity.
            paths.Add(new ActionObjectPath
            {
                ObjectPath = new Identity
                {
                    Id = webIdentityId,
                    Name = CsomIdentity.Web(siteId, webId)
                }
            });

            // The <Property> declaration alone is not enough: CSOM needs an <ObjectPath> action to
            // instantiate it server side before anything can be set on it. Every object path in
            // PnP Core's own GetParentOfTermRequest carries one; omitting it produced a request
            // SharePoint accepted, answered with an empty body, and quietly did nothing.
            paths.Add(new ActionObjectPath
            {
                Action = new BaseAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = regionalSettingsId.ToString()
                },
                ObjectPath = new Property
                {
                    Id = regionalSettingsId,
                    ParentId = webIdentityId,
                    Name = "RegionalSettings"
                }
            });

            foreach (KeyValuePair<string, object> property in properties)
            {
                paths.Add(new ActionObjectPath
                {
                    Action = new SetPropertyAction
                    {
                        Id = idProvider.GetActionId(),
                        ObjectPathId = regionalSettingsId.ToString(),
                        Name = property.Key,
                        SetParameter = new Parameter
                        {
                            Type = PropertyTypes[property.Key],
                            Value = property.Value
                        }
                    }
                });
            }

            // Commit on the settings object - NOT on the web. See the remarks.
            paths.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = regionalSettingsId.ToString(),
                    Name = "Update",
                    Parameters = new List<Parameter>()
                }
            });

            return paths;
        }

        /// <summary>
        /// The raw CSOM response body, kept for diagnosis.
        /// </summary>
        /// <remarks>
        /// <c>ApiCallResponse.Json</c> and <c>StatusCode</c> are <b>not populated on the CSOM
        /// path</b> - PnP Core routes the body here instead. Capturing it is the only way to see
        /// what SharePoint actually said about a request whose actions failed silently.
        /// </remarks>
        internal string RawResponse { get; private set; }

        public void ProcessResponse(string response)
        {
            RawResponse = response;
        }
    }
}
