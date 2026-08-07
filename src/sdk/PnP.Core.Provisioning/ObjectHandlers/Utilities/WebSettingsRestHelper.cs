using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// The two web-level writes PnP Core does not model: supported UI languages
    /// (<b>T2</b>) and regional settings (<b>T1</b>).
    /// </summary>
    /// <remarks>
    /// <para><b>Why these live here rather than in PnP Core.</b> The backlog filed both as "extend
    /// <c>PnP.Core</c>", which would mean changing the core SDK's public surface — an upstream
    /// negotiation for two endpoints only the provisioning engine calls. Both are plain REST, so
    /// doing them from this project through <c>ApiRequest</c> is smaller, self-contained, and
    /// changes nothing anyone else depends on.</para>
    /// <para>If PnP Core later models them properly, these become one-line forwarders and the
    /// handlers do not change. Contributing them upstream stays worthwhile — see
    /// <c>09-open-questions.md</c> — but it is not on this migration's critical path.</para>
    /// </remarks>
    internal static class WebSettingsRestHelper
    {
        #region T1 - regional settings

        /// <summary>
        /// Reads the web's regional settings.
        /// </summary>
        /// <remarks>
        /// <para><b>The read side is a gap too</b>, which the backlog did not record.
        /// <c>IRegionalSettings</c> is not merely read-only - it models a <em>different set of
        /// properties</em>. It exposes display formatting (<c>AM</c>, <c>DateSeparator</c>,
        /// <c>DigitGrouping</c>, …) but not the settings the provisioning schema carries:
        /// <c>CalendarType</c>, <c>AlternateCalendarType</c>, <c>AdjustHijriDays</c>,
        /// <c>Collation</c>, <c>FirstWeekOfYear</c>, <c>WorkDays</c>, <c>WorkDayStartHour</c> and
        /// <c>WorkDayEndHour</c> are all absent.</para>
        /// <para>So extraction reads the raw REST payload rather than the model. Doing both
        /// directions the same way also means read and write cannot drift apart.</para>
        /// </remarks>
        internal static async Task<JsonElement> GetRegionalSettingsAsync(PnPContext context)
        {
            ApiRequestResponse response = await context.Web.ExecuteRequestAsync(
                new ApiRequest(ApiRequestType.SPORest, "_api/web/regionalsettings")).ConfigureAwait(false);

            return JsonDocument.Parse(response.Response).RootElement.Clone();
        }

        /// <summary>
        /// Finds a named property anywhere in the payload.
        /// </summary>
        /// <remarks>
        /// <para><b>Not a top-level lookup.</b> Depending on the OData mode PnP Core negotiates, the
        /// properties may sit at the root or inside a wrapper (<c>d</c>, <c>value</c>, …). Assuming
        /// the root cost real time: every read silently returned its fallback, the extract test
        /// "passed" because its assertions were too weak, and the resulting all-zero model was then
        /// used as the baseline for diagnosing the <em>write</em> - so that investigation was
        /// comparing against garbage.</para>
        /// <para>Searching for the property is a few microseconds and cannot be wrong about shape.</para>
        /// </remarks>
        private static bool TryFind(JsonElement element, string name, out JsonElement found)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty property in element.EnumerateObject())
                {
                    if (property.NameEquals(name))
                    {
                        found = property.Value;
                        return true;
                    }

                    if (TryFind(property.Value, name, out found))
                    {
                        return true;
                    }
                }
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in element.EnumerateArray())
                {
                    if (TryFind(item, name, out found))
                    {
                        return true;
                    }
                }
            }

            found = default;
            return false;
        }

        /// <summary>
        /// Reads an integer from a regional settings payload, falling back when absent.
        /// </summary>
        internal static int GetInt(JsonElement element, string name, int fallback = 0)
        {
            if (!TryFind(element, name, out JsonElement value))
            {
                return fallback;
            }

            // SharePoint returns some of these as JSON strings rather than numbers.
            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int number))
            {
                return number;
            }

            return value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out int parsed)
                ? parsed
                : fallback;
        }

        /// <summary>
        /// Reads a boolean from a regional settings payload, falling back when absent.
        /// </summary>
        internal static bool GetBool(JsonElement element, string name, bool fallback = false)
        {
            if (!TryFind(element, name, out JsonElement value))
            {
                return fallback;
            }

            if (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False)
            {
                return value.GetBoolean();
            }

            return value.ValueKind == JsonValueKind.String && bool.TryParse(value.GetString(), out bool parsed)
                ? parsed
                : fallback;
        }


        /// <summary>
        /// Sets the web's time zone by its SharePoint id.
        /// </summary>
        /// <remarks>
        /// The time zone is not a scalar on <c>SP.RegionalSettings</c> - it is a navigation
        /// property - so it is set through its own endpoint rather than in the MERGE above.
        /// </remarks>
        internal static async Task SetTimeZoneAsync(PnPContext context, int timeZoneId)
        {
            await context.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest,
                $"_api/web/regionalsettings/timezones/getbyid({timeZoneId.ToString(CultureInfo.InvariantCulture)})", null))
                .ConfigureAwait(false);
        }

        #endregion
    }
}
