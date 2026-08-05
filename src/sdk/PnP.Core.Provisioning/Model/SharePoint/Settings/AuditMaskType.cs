namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// Specifies the types of events that are audited on a site collection.
    /// </summary>
    /// <remarks>
    /// Local replacement for <c>Microsoft.SharePoint.Client.AuditMaskType</c>, which the provisioning
    /// model depended on before the move to PnP Core. PnP Core has no equivalent - <c>ISite.Audit</c>
    /// is commented out in the SDK - so the values are reproduced here exactly as CSOM defines them.
    ///
    /// Two details are deliberate and must not be "tidied up":
    ///
    /// 1. There is no <see cref="System.FlagsAttribute"/>, matching CSOM. The values are still combined
    ///    with <c>|=</c> (see the audit flag resolvers), which C# permits on any enum. Adding
    ///    <c>[Flags]</c> would change <c>ToString()</c> to emit a comma separated list of names instead
    ///    of the numeric value, altering how audit settings serialize.
    /// 2. <see cref="All"/> is -1, not the sum of the individual flags.
    ///
    /// Member names are load bearing: the schema serializers resolve audit flags with
    /// <c>Enum.Parse</c> against the names declared in the provisioning schema.
    /// </remarks>
    public enum AuditMaskType
    {
        /// <summary>
        /// No events are audited.
        /// </summary>
        None = 0,

        /// <summary>
        /// Checking out a document.
        /// </summary>
        CheckOut = 1,

        /// <summary>
        /// Checking in a document.
        /// </summary>
        CheckIn = 2,

        /// <summary>
        /// Viewing an item.
        /// </summary>
        View = 4,

        /// <summary>
        /// Deleting or restoring an item.
        /// </summary>
        ObjectDelete = 8,

        /// <summary>
        /// Editing an item.
        /// </summary>
        Update = 16,

        /// <summary>
        /// Changing a user profile.
        /// </summary>
        ProfileChange = 32,

        /// <summary>
        /// Deleting a child item.
        /// </summary>
        ChildDelete = 64,

        /// <summary>
        /// Changing the schema of a content type or column.
        /// </summary>
        SchemaChange = 128,

        /// <summary>
        /// Changing permissions or security settings.
        /// </summary>
        SecurityChange = 256,

        /// <summary>
        /// Restoring an item from the recycle bin.
        /// </summary>
        Undelete = 512,

        /// <summary>
        /// A workflow event.
        /// </summary>
        Workflow = 1024,

        /// <summary>
        /// Copying an item.
        /// </summary>
        Copy = 2048,

        /// <summary>
        /// Moving an item.
        /// </summary>
        Move = 4096,

        /// <summary>
        /// Searching site content.
        /// </summary>
        Search = 8192,

        /// <summary>
        /// All events are audited.
        /// </summary>
        All = -1
    }
}
