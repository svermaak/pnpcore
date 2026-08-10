using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;
using WorkflowsModel = PnP.Core.Provisioning.Model.Workflows;
using WorkflowDefinitionModel = PnP.Core.Provisioning.Model.WorkflowDefinition;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Live coverage for <c>ObjectWorkflows</c> (<b>T16</b>).
    /// </summary>
    /// <remarks>
    /// <para><b>SharePoint 2013 workflows are retired</b>, and the phase-4 request tests already
    /// established that the service is off on this tenant: <c>WorkflowServicesManager</c> comes back
    /// null and every call fails with <i>"Cannot invoke method or retrieve property from null
    /// object"</i>.</para>
    /// <para>So what is testable here is not the workflow behaviour but the <b>handler's</b>
    /// behaviour on a tenant without the service - which is the case almost every consumer will hit.
    /// A template carrying a workflow element must not fail the whole run, and must say once, in
    /// terms a reader can act on, that the platform is gone.</para>
    /// </remarks>
    [TestClass]
    public class ObjectWorkflowsLiveTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Workflows_ReportOnceAndCarryOnWhenTheServiceIsRetired()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                int warnings = 0;
                string lastWarning = null;

                // Assigned explicitly: ProvisioningTemplate.Workflows is null until something sets
                // it, unlike most of the template's collections. A deserialized template gets one
                // from the serializer; a hand-built one has to make it.
                var template = new ProvisioningTemplate
                {
                    Workflows = new WorkflowsModel(),
                };

                // Three definitions on purpose: the handler probes the service once and stops, so a
                // regression that dropped the probe would report three times instead of one.
                for (int i = 1; i <= 3; i++)
                {
                    template.Workflows.WorkflowDefinitions.Add(new WorkflowDefinitionModel
                    {
                        Id = Guid.NewGuid(),
                        DisplayName = $"{TestPrefix}Workflow{i}",
                        XamlPath = $"{TestPrefix}Workflow{i}.xaml",
                        Published = true,
                    });
                }

                var configuration = new ApplyConfiguration
                {
                    MessagesDelegate = (message, type) =>
                    {
                        Console.WriteLine($"[{type}] {message}");

                        if (type == ProvisioningMessageType.Warning)
                        {
                            warnings++;
                            lastWarning = message;
                        }
                    },
                };

                // Must not throw: a retired platform is a tenant state, not a broken template.
                await context.GetProvisioningManager().ApplyTemplateAsync(template, configuration).ConfigureAwait(false);

                Console.WriteLine($"{warnings} warning(s)");

                if (lastWarning == null)
                {
                    Assert.Inconclusive("The workflow service appears to be available on this tenant, " +
                        "so the retired-platform path could not be exercised.");
                }

                // One warning, not one per definition. Reporting the same missing service three
                // times buries the single fact that matters.
                Assert.AreEqual(1, warnings,
                    "The handler reported the missing workflow service more than once.");

                StringAssert.Contains(lastWarning, "retired",
                    "The warning does not tell the reader that the platform is retired.");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Workflows_ExtractIsSilentWhenTheServiceIsRetired()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                var configuration = new ExtractConfiguration();
                configuration.Handlers.Add(ConfigurationHandler.Workflows);

                ProvisioningTemplate extracted = await context.GetProvisioningManager()
                    .GetTemplateAsync(configuration).ConfigureAwait(false);

                Console.WriteLine($"Workflows element: {(extracted.Workflows == null ? "(none)" : "present")}");

                // No element at all is the right answer on a tenant with no workflow service - and
                // it is also what a site with no workflows should produce. An empty <pnp:Workflows/>
                // in every extracted template would be noise, and the element is null by default
                // precisely so that "absent" is expressible.
                if (extracted.Workflows != null)
                {
                    Assert.AreEqual(0, extracted.Workflows.WorkflowDefinitions.Count,
                        "Workflow definitions were extracted from a tenant with no workflow service.");
                }
            }
        }
    }
}
