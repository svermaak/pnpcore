using PnP.Core.Test.Common.Utilities;
using System;

namespace PnP.Core.Provisioning.Test.Utilities
{
    public sealed class TestCommon : TestCommonBase
    {
        private static readonly Lazy<TestCommon> _lazyInstance = new Lazy<TestCommon>(() => new TestCommon(), true);

        /// <summary>
        /// Gets the single TestCommon instance, singleton pattern
        /// </summary>
        internal static TestCommon Instance
        {
            get
            {
                return _lazyInstance.Value;
            }
        }

        /// <summary>
        /// Private constructor since this is a singleton
        /// </summary>
        private TestCommon()
        {
            // The provisioning engine is verified against a real tenant: its failure modes (artefact ordering,
            // throttling, eventual consistency, feature dependencies) are exactly what recorded responses
            // cannot reproduce. See migration doc 07-test-strategy.md, decision D4.
            //
            // Offline tests (model + serializer round-trips) do not create a PnPContext at all and are
            // unaffected by this setting.
            Mocking = false;
        }
    }
}
