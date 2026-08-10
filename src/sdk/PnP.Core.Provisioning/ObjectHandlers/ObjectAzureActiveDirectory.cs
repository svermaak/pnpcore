using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Connectors;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Provisioning.Utilities;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using UserModel = PnP.Core.Provisioning.Model.AzureActiveDirectory.User;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Creates the Microsoft Entra users a tenant template declares, and assigns their licences.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T12</b>, phase 8. Entirely Microsoft Graph, so nothing had to be added to PnP
    /// Core - <c>ExecuteRequestAsync</c> with <c>ApiRequestType.Graph</c> covers it, and PnP Core's
    /// authentication supplies the token that PnP Framework had to acquire by hand for every single
    /// request.</para>
    /// <para><b>Create or update, not create.</b> A template is routinely re-applied, and a user
    /// principal name that already exists answers <c>Request_ResourceExists</c>. The existing user is
    /// found and patched instead - which is also how a template updates a user's details.</para>
    /// <para><b>Licences are a delta, not a list.</b> Graph's <c>assignLicense</c> takes what to add
    /// and what to remove, so anything the user currently holds that the template does not name has
    /// to be worked out and removed explicitly. Sending only additions leaves stale licences
    /// assigned.</para>
    /// </remarks>
    internal class ObjectAzureActiveDirectory : ObjectHierarchyHandlerBase
    {
        public override string Name => "Azure Active Directory";

        public override bool WillProvision(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ApplyConfiguration configuration)
        {
            _willProvision ??= hierarchy?.AzureActiveDirectory?.Users?.Count > 0;
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ExtractConfiguration configuration)
        {
            // Matching PnP Framework: users are not extracted. A tenant's directory is not a property
            // of the sites being extracted, and a template full of real people is not something to
            // produce by accident.
            _willExtract ??= false;
            return _willExtract.Value;
        }

        public override Task<ProvisioningHierarchy> ExtractObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            ExtractConfiguration configuration)
        {
            return Task.FromResult(hierarchy);
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            if (!(hierarchy?.AzureActiveDirectory?.Users?.Count > 0))
            {
                return parser;
            }

            int index = 0;

            foreach (UserModel user in hierarchy.AzureActiveDirectory.Users)
            {
                index++;

                string principalName = parser.ParseString(user.UserPrincipalName);
                WriteSubProgress("User", principalName, index, hierarchy.AzureActiveDirectory.Users.Count);

                try
                {
                    string userId = await CreateOrUpdateUserAsync(context, user, principalName, parser)
                        .ConfigureAwait(false);

                    if (string.IsNullOrEmpty(userId))
                    {
                        continue;
                    }

                    if (user.Licenses?.Count > 0)
                    {
                        await AssignLicensesAsync(context, userId, user, parser).ConfigureAwait(false);
                    }

                    await SetPhotoAsync(context, hierarchy.Connector, user, userId, parser).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    string warning = $"The user '{principalName}' could not be provisioned: {ErrorText.Describe(ex)}";
                    context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                }
            }

            WriteMessage("Done processing users", ProvisioningMessageType.Completed);

            return parser;
        }

        #region Users

        /// <summary>
        /// Creates the user, or patches the one that already has that principal name.
        /// </summary>
        /// <returns>The user's object id, or null if it could not be created or found.</returns>
        private async Task<string> CreateOrUpdateUserAsync(PnPContext context, UserModel user, string principalName,
            TokenParser parser)
        {
            Dictionary<string, object> body = BuildUserBody(user, principalName, parser);

            try
            {
                ApiRequestResponse response = await context.Web.ExecuteRequestAsync(new ApiRequest(
                    HttpMethod.Post, ApiRequestType.Graph, "users", JsonSerializer.Serialize(body)))
                    .ConfigureAwait(false);

                return IdOf(response.Response);
            }
            catch (Exception ex) when (IsAlreadyExists(ex))
            {
                string existingId = await FindUserIdAsync(context, principalName).ConfigureAwait(false);

                if (existingId == null)
                {
                    // Graph said the principal name is taken and then could not find it, which is
                    // worth reporting rather than swallowing - it usually means a soft-deleted user.
                    string warning = $"'{principalName}' already exists but could not be read back, so it " +
                        "was skipped. A soft deleted user with the same principal name will do this.";
                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                    return null;
                }

                // A password profile cannot be patched onto an existing user the way it is set on a
                // new one, and re-sending it would reset a real person's password.
                body.Remove("passwordProfile");
                body.Remove("userPrincipalName");

                await context.Web.ExecuteRequestAsync(new ApiRequest(
                    new HttpMethod("PATCH"), ApiRequestType.Graph, $"users/{existingId}",
                    JsonSerializer.Serialize(body))).ConfigureAwait(false);

                return existingId;
            }
        }

        private static Dictionary<string, object> BuildUserBody(UserModel user, string principalName, TokenParser parser)
        {
            var body = new Dictionary<string, object>
            {
                ["accountEnabled"] = user.AccountEnabled,
                ["displayName"] = parser.ParseString(user.DisplayName),
                ["mailNickname"] = parser.ParseString(user.MailNickname),
                ["userPrincipalName"] = principalName,
                ["givenName"] = parser.ParseString(user.GivenName),
                ["surname"] = parser.ParseString(user.Surname),
                ["jobTitle"] = parser.ParseString(user.JobTitle),
                ["mobilePhone"] = parser.ParseString(user.MobilePhone),
                ["officeLocation"] = parser.ParseString(user.OfficeLocation),
                ["preferredLanguage"] = parser.ParseString(user.PreferredLanguage),
                ["userType"] = "Member",
                ["usageLocation"] = parser.ParseString(user.UsageLocation),
                ["passwordPolicies"] = parser.ParseString(user.PasswordPolicies),
            };

            if (user.PasswordProfile != null)
            {
                body["passwordProfile"] = new Dictionary<string, object>
                {
                    ["forceChangePasswordNextSignIn"] = user.PasswordProfile.ForceChangePasswordNextSignIn,
                    ["forceChangePasswordNextSignInWithMfa"] = user.PasswordProfile.ForceChangePasswordNextSignInWithMfa,
                    ["password"] = EncryptionUtility.ToInsecureString(user.PasswordProfile.Password),
                };
            }

            // Graph rejects an explicit null for several of these, and a template that omits a field
            // means "leave it alone" rather than "clear it".
            foreach (string empty in body.Where(p => p.Value is string s && string.IsNullOrEmpty(s))
                .Select(p => p.Key).ToList())
            {
                body.Remove(empty);
            }

            return body;
        }

        /// <summary>
        /// Finds a user by principal name.
        /// </summary>
        /// <remarks>
        /// Filtered rather than addressed by <c>users/{upn}</c>: that form works, but it answers 404
        /// for a guest whose principal name has been mangled, and the filter is the same round trip.
        /// </remarks>
        private static async Task<string> FindUserIdAsync(PnPContext context, string principalName)
        {
            string filter = Uri.EscapeDataString($"userPrincipalName eq '{principalName.Replace("'", "''")}'");

            ApiRequestResponse response = await context.Web.ExecuteRequestAsync(new ApiRequest(
                ApiRequestType.Graph, $"users?$filter={filter}&$select=id")).ConfigureAwait(false);

            if (string.IsNullOrEmpty(response.Response))
            {
                return null;
            }

            using (JsonDocument document = JsonDocument.Parse(response.Response))
            {
                if (!document.RootElement.TryGetProperty("value", out JsonElement value)
                    || value.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                foreach (JsonElement item in value.EnumerateArray())
                {
                    if (item.TryGetProperty("id", out JsonElement id) && id.ValueKind == JsonValueKind.String)
                    {
                        return id.GetString();
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Whether the failure is Graph saying the object is already there.
        /// </summary>
        /// <remarks>
        /// Matched on the error code in the body rather than on the status: a 400 from
        /// <c>Request_ResourceExists</c> and a 400 from a malformed password policy look identical
        /// otherwise, and treating the second as "already exists" would patch over a real error.
        /// </remarks>
        private static bool IsAlreadyExists(Exception ex)
        {
            string text = ex.ToString();

            return text.IndexOf("Request_ResourceExists", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("ObjectConflict", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string IdOf(string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                return null;
            }

            using (JsonDocument document = JsonDocument.Parse(json))
            {
                return document.RootElement.ValueKind == JsonValueKind.Object
                    && document.RootElement.TryGetProperty("id", out JsonElement id)
                    && id.ValueKind == JsonValueKind.String
                    ? id.GetString()
                    : null;
            }
        }

        #endregion

        #region Licences

        /// <summary>
        /// Brings the user's assigned licences in line with the template.
        /// </summary>
        /// <remarks>
        /// <b>Removals are computed, not declared.</b> <c>assignLicense</c> is a delta operation, so
        /// the licences the user holds have to be read first and everything the template does not
        /// name added to <c>removeLicenses</c>. Sending additions alone leaves a licence the template
        /// dropped still assigned, which reads as success and costs money.
        /// </remarks>
        private async Task AssignLicensesAsync(PnPContext context, string userId, UserModel user, TokenParser parser)
        {
            var wanted = new List<Dictionary<string, object>>();
            var wantedSkus = new HashSet<Guid>();

            foreach (Model.AzureActiveDirectory.UserLicense license in user.Licenses)
            {
                if (!Guid.TryParse(parser.ParseString(license.SkuId), out Guid skuId))
                {
                    Warn(context, $"'{license.SkuId}' is not a usable licence sku id, so it was skipped.");
                    continue;
                }

                wantedSkus.Add(skuId);

                wanted.Add(new Dictionary<string, object>
                {
                    ["skuId"] = skuId,
                    ["disabledPlans"] = (license.DisabledPlans ?? Array.Empty<string>())
                        .Select(p => parser.ParseString(p))
                        .Where(p => Guid.TryParse(p, out _))
                        .Select(Guid.Parse)
                        .ToList(),
                });
            }

            if (wanted.Count == 0)
            {
                return;
            }

            List<Guid> toRemove = (await AssignedSkusAsync(context, userId).ConfigureAwait(false))
                .Where(sku => !wantedSkus.Contains(sku))
                .ToList();

            var body = new Dictionary<string, object>
            {
                ["addLicenses"] = wanted,
                ["removeLicenses"] = toRemove,
            };

            await context.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.Graph,
                $"users/{userId}/assignLicense", JsonSerializer.Serialize(body))).ConfigureAwait(false);
        }

        private static async Task<List<Guid>> AssignedSkusAsync(PnPContext context, string userId)
        {
            var skus = new List<Guid>();

            ApiRequestResponse response = await context.Web.ExecuteRequestAsync(new ApiRequest(
                ApiRequestType.Graph, $"users/{userId}?$select=assignedLicenses")).ConfigureAwait(false);

            if (string.IsNullOrEmpty(response.Response))
            {
                return skus;
            }

            using (JsonDocument document = JsonDocument.Parse(response.Response))
            {
                if (!document.RootElement.TryGetProperty("assignedLicenses", out JsonElement assigned)
                    || assigned.ValueKind != JsonValueKind.Array)
                {
                    return skus;
                }

                foreach (JsonElement license in assigned.EnumerateArray())
                {
                    if (license.TryGetProperty("skuId", out JsonElement skuId)
                        && skuId.ValueKind == JsonValueKind.String
                        && Guid.TryParse(skuId.GetString(), out Guid parsed))
                    {
                        skus.Add(parsed);
                    }
                }
            }

            return skus;
        }

        #endregion

        #region Photo

        /// <summary>
        /// Reports a declared profile photo, which this engine cannot yet upload.
        /// </summary>
        /// <remarks>
        /// <para><b>Refused rather than attempted, on purpose.</b> A photo is a binary PUT to
        /// <c>users/{id}/photo/$value</c>, and PnP Core's <c>ApiRequest.Body</c> is a <c>string</c> -
        /// there is no way to hand it raw bytes. Base64 encoding the image would be accepted by Graph
        /// and store the base64 <em>text</em> as the photo, which is worse than not trying: the run
        /// reports success and the user ends up with a corrupt image.</para>
        /// <para>PnP Framework has its own photo code <b>commented out</b>, with the note that a
        /// photo cannot be set without an already existing mailbox - so nothing is lost against
        /// parity here. What is gained is that a template declaring a photo now says why it did not
        /// get one.</para>
        /// <para>Closing this needs a binary-capable request on <c>PnP.Core</c>; it is filed under
        /// T12 in the porting backlog.</para>
        /// </remarks>
        private Task SetPhotoAsync(PnPContext context, FileConnectorBase connector, UserModel user,
            string userId, TokenParser parser)
        {
            if (string.IsNullOrEmpty(user.ProfilePhoto))
            {
                return Task.CompletedTask;
            }

            Warn(context, $"A profile photo is declared for '{parser.ParseString(user.UserPrincipalName)}' " +
                "but this engine cannot upload one yet: the Graph call needs a binary request body, and " +
                "PnP Core's ApiRequest only carries text. Set it by hand, or through Graph directly.");

            return Task.CompletedTask;
        }

        #endregion

        private void Warn(PnPContext context, string message)
        {
            context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
            WriteMessage(message, ProvisioningMessageType.Warning);
        }
    }
}
