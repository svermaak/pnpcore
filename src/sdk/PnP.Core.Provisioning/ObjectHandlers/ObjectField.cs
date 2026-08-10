using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using FieldModel = PnP.Core.Provisioning.Model.Field;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Creates and updates the site columns a template declares.
    /// </summary>
    /// <remarks>
    /// <para><b>A field is provisioned as XML, not as properties.</b> The template carries the
    /// field's whole <c>SchemaXml</c> and <c>AddFieldAsXmlAsync</c> takes it verbatim. That is why
    /// this handler is far smaller than the surface it covers: every field type SharePoint has ever
    /// had is supported without this code knowing any of them.</para>
    /// <para><b>It runs twice per apply</b>, once for
    /// <see cref="FieldAndListProvisioningStepHelper.Step.ListAndStandardFields"/> and once for
    /// <see cref="FieldAndListProvisioningStepHelper.Step.LookupFields"/>. A lookup field cannot be
    /// created before the list it points at, and a template routinely defines both - so lookups and
    /// calculated fields are held back to the second pass. Changing that ordering breaks templates
    /// that work today; see <c>ProvisioningManager.BuildApplyHandlers</c>.</para>
    /// <para><b>Sub sites are skipped by default.</b> Site columns belong to the root web, and
    /// provisioning them lower down creates a second, unrelated column of the same name. A caller
    /// who means it sets <c>ProvisionFieldsToSubWebs</c>.</para>
    /// </remarks>
    internal class ObjectField : ObjectHandlerBase
    {
        private readonly FieldAndListProvisioningStepHelper.Step step;

        public ObjectField(FieldAndListProvisioningStepHelper.Step step)
        {
            this.step = step;
        }

        public override string Name => $"Fields ({step})";

        public override string InternalName => "Fields";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            _willProvision ??= template.SiteFields.Any();
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            return true;
        }

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl, w => w.Url).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl).ConfigureAwait(false);

                bool provisionToSubWebs = configuration?.ToApplyingInformation()?.ProvisionFieldsToSubWebs ?? false;

                if (IsSubSite(web) && !provisionToSubWebs)
                {
                    const string message = "This template contains site columns and the target is a sub site. " +
                        "Set ProvisionFieldsToSubWebs if you really want them created there.";

                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                    WriteMessage(message, ProvisioningMessageType.Warning);

                    return parser;
                }

                // Everything the update path needs is loaded here, not lazily per field. PnP Core
                // materialises a collection's items once: asking for more properties later leaves
                // the already-loaded items as they were, and reading one then throws "property was
                // not yet loaded" - which reads like a missing await rather than a caching rule.
                await context.Web.LoadAsync(w => w.Fields.QueryProperties(
                    f => f.Id, f => f.SchemaXml, f => f.TypeAsString, f => f.InternalName, f => f.Title, f => f.DefaultValue)).ConfigureAwait(false);

                var existingFieldIds = new HashSet<Guid>(context.Web.Fields.AsRequested().Select(f => f.Id));

                List<FieldModel> fields = OrderFieldsForThisStep(template, parser);

                int currentFieldIndex = 0;
                foreach (FieldModel field in fields)
                {
                    currentFieldIndex++;

                    XElement schema = XElement.Parse(parser.ParseXmlString(field.SchemaXml));
                    string fieldId = schema.Attribute("ID")?.Value;

                    if (!Guid.TryParse(fieldId, out Guid id))
                    {
                        string message = $"A site column in the template has no valid ID attribute and was skipped: {fieldId ?? "<missing>"}";
                        context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                        WriteMessage(message, ProvisioningMessageType.Warning);
                        continue;
                    }

                    string internalName = schema.Attribute("InternalName")?.Value ?? schema.Attribute("Name")?.Value;
                    WriteSubProgress("Field", !string.IsNullOrWhiteSpace(internalName) ? internalName : fieldId, currentFieldIndex, fields.Count);

                    if (existingFieldIds.Contains(id))
                    {
                        await UpdateFieldAsync(context, id, schema, parser, field.SchemaXml).ConfigureAwait(false);
                    }
                    else
                    {
                        await CreateFieldAsync(context, schema, parser, field.SchemaXml).ConfigureAwait(false);
                    }
                }

                WriteMessage("Done processing fields", ProvisioningMessageType.Completed);

                return parser;
            }
        }

        /// <summary>
        /// The template's fields that belong to this pass, ordered by the field they reference.
        /// </summary>
        /// <remarks>
        /// <para>Ordering by the <c>FieldRef</c> attribute is PnP Framework's, and it is preserved
        /// because it is load bearing for dependent field types: a field that names another in
        /// <c>FieldRef</c> has to be created after it.</para>
        /// <para>PnP Framework used a <c>SortedList</c> with a comparer that never returns 0, so
        /// equal keys are kept rather than throwing - a stable-enough sort built out of a container
        /// that does not allow duplicates. Grouping does the same thing without the trick.</para>
        /// </remarks>
        private List<FieldModel> OrderFieldsForThisStep(ProvisioningTemplate template, TokenParser parser)
        {
            var forThisStep = new List<(string FieldRef, FieldModel Field)>();

            foreach (FieldModel siteField in template.SiteFields)
            {
                if (siteField.GetFieldProvisioningStep(parser) != step)
                {
                    continue;
                }

                string fieldRef = (string)XElement.Parse(parser.ParseXmlString(siteField.SchemaXml)).Attribute("FieldRef") ?? string.Empty;
                forThisStep.Add((fieldRef, siteField));
            }

            return forThisStep
                .OrderBy(f => f.FieldRef, StringComparer.Ordinal)
                .Select(f => f.Field)
                .ToList();
        }

        private async Task CreateFieldAsync(PnPContext context, XElement schema, TokenParser parser, string originalFieldXml)
        {
            string fieldXml = parser.ParseXmlString(schema.ToString());

            if (!await IsFieldXmlValidAsync(fieldXml, parser, context).ConfigureAwait(false))
            {
                ThrowInvalidFieldXml(context, parser, fieldXml);
            }

            string internalName = schema.Attribute("Name")?.Value ?? schema.Attribute("StaticName")?.Value;

            fieldXml = await FieldUtilities.FixLookupFieldAsync(context, fieldXml, unresolved =>
            {
                string message = $"The lookup column '{internalName}' points at '{unresolved}', which does not exist on this site. " +
                    "SharePoint will create the column anyway and it will not work.";

                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                WriteMessage(message, ProvisioningMessageType.Warning);
            }).ConfigureAwait(false);

            context.Logger?.LogDebug("{Source}: adding site column {Field}.", Constants.LOGGING_SOURCE, schema.Attribute("ID")?.Value);

            // AddFieldInternalNameHint: without it SharePoint derives the internal name from the
            // display name, so a localized template produces columns whose internal names differ per
            // language - and every reference to them by name then breaks.
            IField field = await context.Web.Fields.AddFieldAsXmlAsync(fieldXml, false, AddFieldOptionsFlags.AddFieldInternalNameHint).ConfigureAwait(false);

            await field.LoadAsync(f => f.Id, f => f.TypeAsString, f => f.DefaultValue, f => f.InternalName, f => f.Title).ConfigureAwait(false);

            // Register the new field's tokens straight away, so a later field in the same template
            // can reference it - a calculated column over a column defined two lines above it.
            parser.AddToken(new FieldTitleToken(context, field.InternalName, field.Title));
            parser.AddToken(new FieldIdToken(context, field.InternalName, field.Id));

            await LocalizeAsync(context, field, originalFieldXml, parser).ConfigureAwait(false);
            await WarnOnTaxonomyDefaultValueAsync(context, field).ConfigureAwait(false);
        }

        private async Task UpdateFieldAsync(PnPContext context, Guid fieldId, XElement templateSchema, TokenParser parser, string originalFieldXml)
        {
            // IFieldCollection has no GetById; the whole collection is already loaded with its ids
            // by the caller, so this is a lookup rather than a request.
            IField existingField = context.Web.Fields.AsRequested().FirstOrDefault(f => f.Id == fieldId);
            if (existingField == null)
            {
                return;
            }

            XElement existingSchema = XElement.Parse(existingField.SchemaXml);

            var comparer = new XNodeEqualityComparer();
            if (comparer.GetHashCode(existingSchema) == comparer.GetHashCode(templateSchema))
            {
                return;
            }

            string existingType = existingSchema.Attribute("Type")?.Value;
            string templateType = templateSchema.Attribute("Type")?.Value;

            if (!string.Equals(existingType, templateType, StringComparison.Ordinal))
            {
                // Changing a field's type in place is not possible, and deleting and re-creating it
                // would take every list item's data with it.
                string fieldName = existingSchema.Attribute("Name")?.Value ?? existingSchema.Attribute("StaticName")?.Value;
                string message = $"The site column '{fieldName}' ({fieldId}) exists but is of type '{existingType}' rather than '{templateType}', so it was skipped.";

                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                WriteMessage(message, ProvisioningMessageType.Warning);

                return;
            }

            if (!await IsFieldXmlValidAsync(parser.ParseXmlString(originalFieldXml), parser, context).ConfigureAwait(false))
            {
                ThrowInvalidFieldXml(context, parser, originalFieldXml);
            }

            // A lookup's List attribute is dropped before merging: it is an id on the existing field
            // and may be a url in the template, and overwriting one form with the other produces a
            // schema SharePoint rejects.
            templateSchema.Attribute("List")?.Remove();

            MergeInto(existingSchema, templateSchema);

            if (string.Equals(templateType, "Calculated", StringComparison.OrdinalIgnoreCase))
            {
                // A calculated field's FieldRefs are derived from its formula. Leaving the old ones
                // in place next to a new formula makes the field fail to update.
                existingSchema.Descendants("FieldRefs").FirstOrDefault()?.Remove();
            }

            // Version is SharePoint's own counter; sending it back rejects the update.
            existingSchema.Attributes("Version").Remove();

            existingField.SchemaXml = parser.ParseXmlString(existingSchema.ToString());

            try
            {
                await existingField.UpdateAndPushChangesAsync().ConfigureAwait(false);
            }
            catch (Exception ex) when (IsTimeout(ex))
            {
                // Pushing a change down to every list that uses a column can outrun the request
                // timeout on a large site. The work continues server side, so this is reported and
                // the run goes on rather than failing a template over a slow column.
                string fieldName = existingSchema.Attribute("Name")?.Value ?? existingSchema.Attribute("StaticName")?.Value;
                string message = $"Updating the site column '{fieldName}' timed out. SharePoint is still applying it; the run continued.";

                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                WriteMessage(message, ProvisioningMessageType.Warning);
            }

            await LocalizeAsync(context, existingField, originalFieldXml, parser).ConfigureAwait(false);

            if (IsTaxonomyField(existingField))
            {
                await WarnOnTaxonomyDefaultValueAsync(context, existingField).ConfigureAwait(false);
                await UpdateTaxonomyFieldAsync(context, existingField, existingSchema).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Copies the template's attributes and elements onto the existing field's schema.
        /// </summary>
        /// <remarks>
        /// A merge rather than a replace: the existing schema carries server-generated attributes the
        /// template does not know about, and dropping them makes SharePoint reject the update.
        /// </remarks>
        private static void MergeInto(XElement existingSchema, XElement templateSchema)
        {
            foreach (XAttribute attribute in templateSchema.Attributes())
            {
                XAttribute existing = existingSchema.Attribute(attribute.Name);
                if (existing != null)
                {
                    existing.Value = attribute.Value;
                }
                else
                {
                    existingSchema.Add(attribute);
                }
            }

            foreach (XElement element in templateSchema.Elements())
            {
                existingSchema.Element(element.Name)?.Remove();
                existingSchema.Add(element);
            }
        }

        /// <summary>
        /// Writes the per-language display name and description, when the template uses tokens.
        /// </summary>
        private async Task LocalizeAsync(PnPContext context, IField field, string originalFieldXml, TokenParser parser)
        {
            if (!UserResources.ContainsResourceToken(originalFieldXml))
            {
                return;
            }

            XElement original = XElement.Parse(originalFieldXml);
            string displayName = original.Attribute("DisplayName")?.Value ?? string.Empty;
            string description = original.Attribute("Description")?.Value ?? string.Empty;

            if (!UserResources.ContainsResourceToken(displayName) && !UserResources.ContainsResourceToken(description))
            {
                return;
            }

            (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

            if (UserResources.ContainsResourceToken(displayName))
            {
                await UserResources.TrySetAsync(context,
                    UserResourcePath.ForField(siteId, webId, field.Id, ResourceProperty.Title), displayName, parser,
                    $"the display name of site column '{field.InternalName}'",
                    m => WriteMessage(m, ProvisioningMessageType.Warning)).ConfigureAwait(false);
            }

            if (UserResources.ContainsResourceToken(description))
            {
                await UserResources.TrySetAsync(context,
                    UserResourcePath.ForField(siteId, webId, field.Id, ResourceProperty.Description), description, parser,
                    $"the description of site column '{field.InternalName}'",
                    m => WriteMessage(m, ProvisioningMessageType.Warning)).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Applies the taxonomy-specific settings a field's schema carries.
        /// </summary>
        /// <remarks>
        /// These live in the field's <c>Customization</c> block rather than as schema attributes, so
        /// the merge above does not reach them - they have to be set as properties on the field.
        /// </remarks>
        private static async Task UpdateTaxonomyFieldAsync(PnPContext context, IField field, XElement schema)
        {
            await field.LoadAsync(f => f.SspId, f => f.TermSetId, f => f.AnchorId, f => f.Open, f => f.IsPathRendered).ConfigureAwait(false);

            bool dirty = false;

            dirty |= SetGuidIfChanged(schema, "SspId", field.SspId, v => field.SspId = v);
            dirty |= SetGuidIfChanged(schema, "TermSetId", field.TermSetId, v => field.TermSetId = v);
            dirty |= SetGuidIfChanged(schema, "AnchorId", field.AnchorId, v => field.AnchorId = v);
            dirty |= SetBoolIfChanged(schema, "Open", field.Open, v => field.Open = v);
            dirty |= SetBoolIfChanged(schema, "IsPathRendered", field.IsPathRendered, v => field.IsPathRendered = v);

            if (dirty)
            {
                await field.UpdateAndPushChangesAsync().ConfigureAwait(false);
                context.Logger?.LogDebug("{Source}: updated the taxonomy settings of {Field}.", Constants.LOGGING_SOURCE, field.InternalName);
            }
        }

        /// <summary>
        /// Reports a taxonomy default value that cannot yet be applied correctly.
        /// </summary>
        /// <remarks>
        /// <b>Backlog T10.</b> A taxonomy default is stored as <c>WssId;#Label|TermGuid</c>, and the
        /// <c>WssId</c> is a per-site lookup id that means nothing on the target. CSOM resolved this
        /// with <c>TaxonomyField.GetValidatedString</c>; PnP Core has no equivalent, and no CSOM
        /// request for it exists yet. Written as-is the default is accepted and then does not work -
        /// so it is reported rather than left to be discovered by a user.
        /// </remarks>
        private Task WarnOnTaxonomyDefaultValueAsync(PnPContext context, IField field)
        {
            if (!IsTaxonomyField(field))
            {
                return Task.CompletedTask;
            }

            string defaultValue = field.DefaultValue?.ToString();
            if (string.IsNullOrEmpty(defaultValue))
            {
                return Task.CompletedTask;
            }

            string message = string.Format(CultureInfo.CurrentCulture,
                "The taxonomy column '{0}' has a default value ('{1}'). Its WssId has to be re-resolved against this site, " +
                "which needs GetValidatedString - backlog T10, later in phase 6. The default was written as it stands and may not resolve.",
                field.InternalName, defaultValue);

            context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
            WriteMessage(message, ProvisioningMessageType.Warning);

            return Task.CompletedTask;
        }

        #endregion

        #region Extract

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl, w => w.Url, w => w.Id).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl).ConfigureAwait(false);

                if (IsSubSite(web))
                {
                    // Site columns belong to the root web; a sub site's copies are the same columns.
                    context.Logger?.LogDebug("{Source}: this is a sub site, so no site columns were extracted.", Constants.LOGGING_SOURCE);
                    return template;
                }

                // TextField comes along because a taxonomy column's hidden text column has to be
                // moved to the front of the template - and asking for it per field afterwards hits
                // PnP Core's "materialised once" rule and throws.
                await context.Web.LoadAsync(
                    w => w.Fields.QueryProperties(f => f.Id, f => f.SchemaXml, f => f.TypeAsString,
                        f => f.InternalName, f => f.Title, f => f.TextField),
                    w => w.Lists.QueryProperties(l => l.Id, l => l.Title)).ConfigureAwait(false);

                ProvisioningTemplateCreationInformation creationInformation = configuration?.ToCreationInformation();

                var taxonomyTextFieldsToMoveUp = new List<Guid>();
                var calculatedFieldsToMoveDown = new List<Guid>();

                List<IField> fields = context.Web.Fields.AsRequested().Where(f => !BuiltInFieldId.Contains(f.Id)).ToList();

                int currentFieldIndex = 0;
                foreach (IField field in fields)
                {
                    currentFieldIndex++;
                    WriteSubProgress("Field", field.InternalName, currentFieldIndex, fields.Count);

                    string fieldXml = await TokenizeFieldAsync(context, web, field, taxonomyTextFieldsToMoveUp, calculatedFieldsToMoveDown).ConfigureAwait(false);

                    fieldXml = await PersistResourcesAsync(context, field, fieldXml, template, creationInformation).ConfigureAwait(false);

                    template.SiteFields.Add(new FieldModel { SchemaXml = fieldXml });
                }

                ReorderForApply(template, taxonomyTextFieldsToMoveUp, calculatedFieldsToMoveDown);

                if (creationInformation?.BaseTemplate != null)
                {
                    RemoveBaseTemplateFields(template, creationInformation.BaseTemplate);
                }

                WriteMessage("Done processing fields", ProvisioningMessageType.Completed);

                return template;
            }
        }

        private async Task<string> TokenizeFieldAsync(PnPContext context, IWeb web, IField field,
            List<Guid> taxonomyTextFieldsToMoveUp, List<Guid> calculatedFieldsToMoveDown)
        {
            string fieldXml = field.SchemaXml;
            XElement element = XElement.Parse(fieldXml);

            // A lookup names its target list by id, which is generated per site.
            if (!string.IsNullOrEmpty(element.Attribute("List")?.Value))
            {
                fieldXml = TokenizeListAndSiteIds(context, web, fieldXml);
                element = XElement.Parse(fieldXml);
            }

            if (field.TypeAsString != null && field.TypeAsString.StartsWith("TaxonomyField", StringComparison.Ordinal))
            {
                taxonomyTextFieldsToMoveUp.Add(field.TextField);

                fieldXml = await TokenizeTaxonomyFieldAsync(context, element).ConfigureAwait(false);
                element = XElement.Parse(fieldXml);
            }

            // Version is SharePoint's counter, not part of the field's definition.
            if (element.Attribute("Version") != null)
            {
                element.Attributes("Version").Remove();
                fieldXml = element.ToString();
            }

            if (element.Attribute("Type")?.Value == "Calculated")
            {
                fieldXml = TokenizeFieldFormula(context, element);
                calculatedFieldsToMoveDown.Add(field.Id);
            }

            return fieldXml;
        }

        /// <summary>
        /// Rewrites a calculated field's formula to reference columns by display name.
        /// </summary>
        /// <remarks>
        /// A stored formula references columns by <em>internal</em> name, but SharePoint only accepts
        /// a formula written with <em>display</em> names when a field is created. The
        /// <c>FieldRefs</c> block is dropped with it: it is derived from the formula, and sending a
        /// stale one back makes the field fail to provision.
        /// </remarks>
        private static string TokenizeFieldFormula(PnPContext context, XElement schema)
        {
            XElement formulaElement = schema.Descendants("Formula").FirstOrDefault();
            if (formulaElement == null)
            {
                return schema.ToString();
            }

            string formula = formulaElement.Value;
            if (!string.IsNullOrEmpty(formula))
            {
                List<IField> allFields = context.Web.Fields.AsRequested().ToList();

                foreach (string internalName in schema.Descendants("FieldRef")
                    .Select(fieldRef => fieldRef.Attribute("Name")?.Value)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Distinct())
                {
                    IField referenced = allFields.FirstOrDefault(f => f.InternalName == internalName);
                    if (referenced != null)
                    {
                        // The stored formula may already bracket the name, so any brackets already
                        // present are consumed rather than nested. PnP Framework replaced the bare
                        // name and produced "=[[Column]]" for every formula SharePoint had stored in
                        // bracketed form - which is most of them, and which will not provision.
                        formula = Regex.Replace(formula, $@"\[?{Regex.Escape(internalName)}\]?", $"[{referenced.Title}]");
                    }
                }

                formulaElement.Value = formula;
            }

            schema.Descendants("FieldRefs").Remove();

            return schema.ToString();
        }

        private static string TokenizeListAndSiteIds(PnPContext context, IWeb web, string schemaXml)
        {
            foreach (IList list in context.Web.Lists.AsRequested())
            {
                schemaXml = Regex.Replace(schemaXml, list.Id.ToString(),
                    $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
            }

            // The braced form is replaced first and with a doubled brace, because the result is fed
            // back through the token parser - which would otherwise read "{{siteid}}" as an escape.
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("B"), "{{siteid}}", RegexOptions.IgnoreCase);
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("D"), "{siteid}", RegexOptions.IgnoreCase);

            return schemaXml;
        }

        /// <summary>
        /// Reads a field's display name and description in every supported language.
        /// </summary>
        private static async Task<string> PersistResourcesAsync(PnPContext context, IField field, string fieldXml,
            ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation)
        {
            if (creationInformation?.PersistMultiLanguageResources != true || template.SupportedUILanguages.Count == 0)
            {
                return fieldXml;
            }

            // A field the base template already provides is dropped from the template later, so
            // reading its translations - one round trip per field - would be wasted work.
            if (creationInformation.BaseTemplate != null
                && creationInformation.BaseTemplate.SiteFields.Any(f => IdOf(f.SchemaXml) == field.Id))
            {
                return fieldXml;
            }

            (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);
            XElement element = XElement.Parse(fieldXml);
            string key = field.Title.Replace(" ", "_");

            string titleToken = $"Field_{key}_DisplayName";
            if (await UserResources.PersistAsync(context,
                UserResourcePath.ForField(siteId, webId, field.Id, ResourceProperty.Title), titleToken, template, creationInformation).ConfigureAwait(false))
            {
                element.SetAttributeValue("DisplayName", UserResources.TokenFor(titleToken));
            }

            string descriptionToken = $"Field_{key}_Description";
            if (await UserResources.PersistAsync(context,
                UserResourcePath.ForField(siteId, webId, field.Id, ResourceProperty.Description), descriptionToken, template, creationInformation).ConfigureAwait(false))
            {
                element.SetAttributeValue("Description", UserResources.TokenFor(descriptionToken));
            }

            return element.ToString();
        }

        /// <summary>
        /// Puts the extracted fields into an order that will apply cleanly.
        /// </summary>
        /// <remarks>
        /// <para>A taxonomy column has a hidden text column behind it that must exist first, so those
        /// go to the front. Calculated columns reference other columns, so they go to the back.</para>
        /// <para><b>One calculated column referencing another still breaks.</b> Moving them all to
        /// the end does not order them among themselves - that needs a dependency sort, and PnP
        /// Framework carried the same limitation with the same comment.</para>
        /// </remarks>
        private static void ReorderForApply(ProvisioningTemplate template, List<Guid> taxonomyTextFieldsToMoveUp, List<Guid> calculatedFieldsToMoveDown)
        {
            foreach (Guid textFieldId in taxonomyTextFieldsToMoveUp)
            {
                FieldModel field = template.SiteFields.FirstOrDefault(f => IdOf(f.SchemaXml) == textFieldId);
                if (field == null)
                {
                    continue;
                }

                template.SiteFields.RemoveAll(f => IdOf(f.SchemaXml) == textFieldId);
                template.SiteFields.Insert(0, field);
            }

            foreach (Guid calculatedFieldId in calculatedFieldsToMoveDown)
            {
                FieldModel field = template.SiteFields.FirstOrDefault(f => IdOf(f.SchemaXml) == calculatedFieldId);
                if (field == null)
                {
                    continue;
                }

                template.SiteFields.RemoveAll(f => IdOf(f.SchemaXml) == calculatedFieldId);
                template.SiteFields.Add(field);
            }
        }

        private static void RemoveBaseTemplateFields(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            foreach (FieldModel field in baseTemplate.SiteFields)
            {
                Guid id = IdOf(field.SchemaXml);
                if (id != Guid.Empty)
                {
                    template.SiteFields.RemoveAll(f => IdOf(f.SchemaXml) == id);
                }
            }
        }

        #endregion

        #region Helpers

        private static Guid IdOf(string schemaXml)
        {
            try
            {
                return Guid.TryParse(XElement.Parse(schemaXml).Attribute("ID")?.Value, out Guid id) ? id : Guid.Empty;
            }
            catch (System.Xml.XmlException)
            {
                return Guid.Empty;
            }
        }

        private static bool IsTaxonomyField(IField field)
        {
            return field.TypeAsString == "TaxonomyFieldType" || field.TypeAsString == "TaxonomyFieldTypeMulti";
        }

        private static bool IsTimeout(Exception ex)
        {
            return ex is TimeoutException
                || ex is TaskCanceledException
                || (ex.Message != null && ex.Message.IndexOf("timeout", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void ThrowInvalidFieldXml(PnPContext context, TokenParser parser, string fieldXml)
        {
            string leftOver = string.Join(" ", parser.GetLeftOverTokens(fieldXml));
            string message = string.IsNullOrWhiteSpace(leftOver)
                ? "A site column's schema XML is not valid - its taxonomy term set could not be resolved on this site."
                : $"A site column's schema XML still contains unresolved tokens:{leftOver}";

            context.Logger?.LogError("{Source}: {Message}", Constants.LOGGING_SOURCE, message);

            throw new InvalidOperationException(message);
        }

        private static bool SetGuidIfChanged(XElement schema, string propertyName, Guid current, Action<Guid> set)
        {
            string raw = TaxonomyProperty(schema, propertyName);

            if (!Guid.TryParse(raw, out Guid value) || value == current)
            {
                return false;
            }

            set(value);
            return true;
        }

        private static bool SetBoolIfChanged(XElement schema, string propertyName, bool current, Action<bool> set)
        {
            string raw = TaxonomyProperty(schema, propertyName);

            if (!bool.TryParse(raw, out bool value) || value == current)
            {
                return false;
            }

            set(value);
            return true;
        }

        /// <summary>
        /// Reads one value out of a taxonomy field's <c>Customization</c> property bag.
        /// </summary>
        private static string TaxonomyProperty(XElement schema, string propertyName)
        {
            return schema.Descendants("Property")
                .FirstOrDefault(p => p.Element("Name")?.Value == propertyName)
                ?.Element("Value")?.Value;
        }

        #endregion
    }
}
