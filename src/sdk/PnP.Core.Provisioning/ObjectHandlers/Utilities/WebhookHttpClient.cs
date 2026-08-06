using System;
using System.Net.Http;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// The <see cref="HttpClient"/> used to deliver provisioning webhooks.
    /// </summary>
    /// <remarks>
    /// <para><b>Deliberately not PnP Core's own client.</b> <c>PnPContext.RestClient.Client</c> is
    /// configured for SharePoint - a SharePoint <c>Accept</c> header, the SDK's client tag, and a
    /// timeout tuned for SharePoint calls. Webhook urls come out of the template and point at
    /// arbitrary third-party endpoints, so they get a plain client with none of that attached.</para>
    /// <para>A single static instance, because that is what <see cref="HttpClient"/> is designed
    /// for: one per call exhausts sockets under load. <see cref="Instance"/> is settable so a host
    /// that manages its own <c>IHttpClientFactory</c> can supply a named client instead.</para>
    /// </remarks>
    public static class WebhookHttpClient
    {
        private static readonly Lazy<HttpClient> defaultClient = new Lazy<HttpClient>(() => new HttpClient(), true);
        private static HttpClient instance;

        /// <summary>
        /// The client webhooks are sent with. Assign your own to route them through an
        /// <c>IHttpClientFactory</c> managed client; assign <c>null</c> to go back to the default.
        /// </summary>
        public static HttpClient Instance
        {
            get => instance ?? defaultClient.Value;
            set => instance = value;
        }
    }
}
