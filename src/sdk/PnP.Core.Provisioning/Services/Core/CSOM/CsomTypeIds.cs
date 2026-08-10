namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// CSOM server type ids - the GUIDs a <c>StaticMethod</c> or <c>Constructor</c> object path is
    /// addressed by.
    /// </summary>
    /// <remarks>
    /// <para>A CSOM request identifies a static method or a constructor by type <em>id</em>, never
    /// by type name. Get one wrong and the request fails at runtime with an unhelpful server error;
    /// nothing catches it at compile time and nothing about the value is guessable.</para>
    /// <para>Every value here was <b>extracted from the shipped CSOM assemblies</b> rather than
    /// copied from documentation, by reading the <c>ServerTypeId</c> named argument of the
    /// <c>[ScriptType]</c> attribute on each type. (The attribute's <em>constructor</em> argument
    /// is the server type <em>name</em>, e.g. <c>SP.Taxonomy.TaxonomySession</c> - a different
    /// thing, and an easy mistake to make.) The extraction was validated against PnP Core's
    /// existing <c>GetParentOfTermRequest</c>, which hardcodes <see cref="TaxonomySession"/>: the
    /// value read back matches it exactly.</para>
    /// <para>Source: <c>microsoft.sharepointonline.csom 16.1.26914.12004</c>, <c>lib/net40-full</c>.
    /// These are stable across CSOM versions - they are part of the wire protocol - but if a new
    /// type is needed, re-run the extraction rather than guessing. The procedure is recorded in
    /// migration doc 08, phase 4.</para>
    /// </remarks>
    internal static class CsomTypeIds
    {
        #region Taxonomy - Microsoft.SharePoint.Client.Taxonomy

        /// <summary>SP.Taxonomy.TaxonomySession</summary>
        internal const string TaxonomySession = "{981cbc68-9edc-4f8d-872f-71146fcbb84f}";

        /// <summary>SP.Taxonomy.TermStore</summary>
        internal const string TermStore = "{9d8a8884-b1dc-4dbc-81c5-ddea8ad3184c}";

        /// <summary>SP.Taxonomy.TermGroup</summary>
        internal const string TermGroup = "{65d76872-0b65-42de-8ebd-d76f6d3491c6}";

        /// <summary>SP.Taxonomy.TermSet</summary>
        internal const string TermSet = "{e26feb13-2940-4db9-a52b-12b160113a80}";

        /// <summary>SP.Taxonomy.TermSetItem - the base of both TermSet and Term, and where
        /// <c>CreateTerm</c>, <c>ReuseTerm</c> and <c>CustomSortOrder</c> live.</summary>
        internal const string TermSetItem = "{a99e4a8f-010b-4e56-9b29-b7bd6ec51263}";

        /// <summary>SP.Taxonomy.Term</summary>
        internal const string Term = "{5b8c81b7-7cd2-40dc-8525-5eca12a4eb73}";

        /// <summary>SP.Taxonomy.Label</summary>
        internal const string Label = "{81503ae1-8747-4684-a172-163c7e009ef9}";

        #endregion

        #region Lists - Microsoft.SharePoint.Client

        /// <summary>
        /// SP.ListCreationInformation - the value object <c>Web.Lists.Add</c> takes.
        /// </summary>
        /// <remarks>
        /// Needed because PnP Core's <c>IListCollection.AddAsync</c> takes only a title and a
        /// template type. A provisioning template also carries the list's <b>Url</b> - which is its
        /// real identity, is routinely different from the title, and is what every
        /// <c>{listurl:…}</c> token and lookup field resolves against - as well as a description,
        /// a quick launch choice and, for custom <c>.stp</c> templates, a feature id.
        /// </remarks>
        internal const string ListCreationInformation = "{e247b7fc-095e-4ea4-a4c9-c5d373723d8c}";

        /// <summary>SP.ViewCreationInformation</summary>
        internal const string ViewCreationInformation = "{a3547807-7266-42f3-b055-afa6e840e458}";

        /// <summary>SP.ListItemCreationInformation</summary>
        internal const string ListItemCreationInformation = "{54cdbee5-0897-44ac-829f-411557fa11be}";

        #endregion

        #region Publishing - Microsoft.SharePoint.Client.Publishing

        /// <summary>SP.Publishing.SiteImageRenditions</summary>
        internal const string SiteImageRenditions = "{324675a4-aa0d-47db-a937-c2e5dc53457e}";

        /// <summary>SP.Publishing.ImageRendition</summary>
        internal const string ImageRendition = "{cb63161f-1f15-446f-9ba9-af89ae03cd45}";

        #endregion

        #region Workflow services - Microsoft.SharePoint.Client.WorkflowServices

        /// <summary>SP.WorkflowServices.WorkflowServicesManager</summary>
        internal const string WorkflowServicesManager = "{4ccc7f0e-bf7e-4477-999c-6458a73d0039}";

        /// <summary>SP.WorkflowServices.WorkflowDeploymentService</summary>
        internal const string WorkflowDeploymentService = "{3573a52f-3a27-4700-a08e-822c191c2c5d}";

        /// <summary>SP.WorkflowServices.WorkflowSubscriptionService</summary>
        internal const string WorkflowSubscriptionService = "{fc956693-2419-4950-8963-52ebc3e46501}";

        /// <summary>SP.WorkflowServices.WorkflowInstanceService</summary>
        internal const string WorkflowInstanceService = "{71252277-2470-4022-bcaf-c4657aa118c3}";

        /// <summary>SP.WorkflowServices.WorkflowDefinition</summary>
        internal const string WorkflowDefinition = "{60320d36-4b4d-4bac-a092-8f8b5610edcd}";

        /// <summary>SP.WorkflowServices.WorkflowSubscription</summary>
        internal const string WorkflowSubscription = "{d185ede6-c3c3-4d37-9e8c-2382deb37708}";

        #endregion

        #region Web parts - Microsoft.SharePoint.Client.WebParts

        /// <summary>SP.WebParts.LimitedWebPartManager</summary>
        internal const string LimitedWebPartManager = "{ac641ade-62df-49c9-af8e-abda6278e920}";

        /// <summary>SP.WebParts.WebPartDefinition</summary>
        internal const string WebPartDefinition = "{44bf1024-6127-432a-8e3d-fb317fb4541e}";

        /// <summary>SP.WebParts.WebPart</summary>
        internal const string WebPart = "{612a6bd9-6c99-43c9-813a-8d7e19702118}";

        #endregion

        #region Core - Microsoft.SharePoint.Client

        /// <summary>SP.Audit</summary>
        internal const string Audit = "{1307502c-2a94-4c1e-8ba6-30da4b0391f1}";

        /// <summary>SP.UserResource</summary>
        internal const string UserResource = "{2b2affeb-3ccd-4996-9864-211c960e647c}";

        /// <summary>SP.Utilities.Utility</summary>
        internal const string Utility = "{16f43e7e-bf35-475d-b677-9dc61e549339}";

        /// <summary>
        /// SP.RegionalSettings.
        /// </summary>
        /// <remarks>
        /// Needed because <b>REST cannot write regional settings at all</b> - see backlog T1.
        /// Note it carries its own <c>Update()</c>, so the change is committed on the settings
        /// object rather than on the web.
        /// </remarks>
        internal const string RegionalSettings = "{84c424a9-a1d6-46ba-8398-c46257ecd25b}";

        /// <summary>SP.TimeZone</summary>
        internal const string TimeZone = "{5519d02c-ce37-4b91-b61d-a1cefe0fc85e}";

        #endregion

        #region Information policy - Microsoft.Office.Client.Policy

        /// <summary>
        /// SP.InformationPolicy.ProjectPolicy.
        /// </summary>
        /// <remarks>
        /// The CSOM type is <c>Microsoft.SharePoint.Client.InformationPolicy.ProjectPolicy</c>, not
        /// <c>Microsoft.Office.RecordsManagement.InformationPolicy.ProjectPolicy</c> as the
        /// server-side API and much of the documentation call it. It ships in
        /// <c>Microsoft.Office.Client.Policy.dll</c>, which matches neither name.
        /// </remarks>
        internal const string ProjectPolicy = "{ec5e0a70-0cc3-408f-a4dc-1bb3495aac75}";

        #endregion

        #region Tenant administration - Microsoft.Online.SharePoint.TenantAdministration

        /// <summary>
        /// Microsoft.Online.SharePoint.TenantAdministration.Tenant.
        /// </summary>
        /// <remarks>
        /// A constructor path, not a static method: every tenant operation hangs off a
        /// <c>new Tenant(context)</c>. The same value is hardcoded in <c>PnP.Core.Admin</c>'s
        /// tenant requests, which is where it was checked against rather than being extracted
        /// again.
        /// </remarks>
        internal const string Tenant = "{268004ae-ef6b-4e9b-8425-127220d84719}";

        #endregion
    }
}
