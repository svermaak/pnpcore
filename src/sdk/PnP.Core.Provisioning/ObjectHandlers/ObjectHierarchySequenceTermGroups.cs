using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;
using TermGroupModel = PnP.Core.Provisioning.Model.TermGroup;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Creates the term groups a tenant template's sequence declares, before its sites are created.
    /// </summary>
    /// <remarks>
    /// <para>Phase 8, and it reuses phase 6's <see cref="TermGroupProvisioner"/> wholesale - the
    /// taxonomy work is identical whether the term groups hang off a template or off a sequence.
    /// Backlog <b>T10</b> is what made that true.</para>
    /// <para><b>The deferred reuses are applied once, after every sequence.</b> A reused term's source
    /// may be declared in a later group, or in a later sequence, so resolving them per group - or per
    /// sequence - would fail on a forward reference. This is the same reason <c>ObjectTermGroups</c>
    /// defers them to the end of the template.</para>
    /// <para><b>Every sequence is processed, not only the one being applied.</b> That is PnP
    /// Framework's behaviour and it is deliberate: term groups are tenant-wide, a sequence's sites
    /// may use terms another sequence defines, and creating a term group that already exists is a
    /// no-op.</para>
    /// </remarks>
    internal class ObjectHierarchySequenceTermGroups : ObjectHierarchyHandlerBase
    {
        public override string Name => "Term Groups";

        public override bool WillProvision(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ApplyConfiguration configuration)
        {
            _willProvision ??= hierarchy?.Sequences?
                .Any(s => s.TermStore?.TermGroups?.Count > 0) == true;

            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ExtractConfiguration configuration)
        {
            // Deferred with the rest of hierarchy extraction - see ObjectHierarchySequenceSites.
            _willExtract ??= false;
            return _willExtract.Value;
        }

        public override Task<ProvisioningHierarchy> ExtractObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            ExtractConfiguration configuration)
        {
            // PnP Framework throws NotImplementedException here. Returning the hierarchy unchanged
            // says the same thing without turning "this part is not built" into a failure of the
            // whole extract.
            return Task.FromResult(hierarchy);
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            if (hierarchy?.Sequences == null)
            {
                return parser;
            }

            var provisioner = new TermGroupProvisioner(context, parser,
                m => WriteMessage(m, ProvisioningMessageType.Warning));

            foreach (ProvisioningSequence sequence in hierarchy.Sequences)
            {
                if (!(sequence.TermStore?.TermGroups?.Count > 0))
                {
                    continue;
                }

                int index = 0;

                foreach (TermGroupModel modelGroup in sequence.TermStore.TermGroups)
                {
                    index++;
                    WriteSubProgress("Term group", modelGroup.Name, index, sequence.TermStore.TermGroups.Count);

                    try
                    {
                        await provisioner.ProcessGroupAsync(modelGroup).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        string warning = $"The term group '{modelGroup.Name}' could not be provisioned: " +
                            ErrorText.Describe(ex);
                        context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                        WriteMessage(warning, ProvisioningMessageType.Warning);
                    }
                }
            }

            await provisioner.ApplyDeferredReusesAsync().ConfigureAwait(false);

            WriteMessage("Done processing term groups", ProvisioningMessageType.Completed);

            return parser;
        }
    }
}
