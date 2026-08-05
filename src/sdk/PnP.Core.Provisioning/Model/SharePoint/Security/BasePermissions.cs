using PnP.Core.Model.SharePoint;
using System;

namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// Represents a set of SharePoint permissions as a 64 bit mask, split into a high and a low word.
    /// </summary>
    /// <remarks>
    /// Local replacement for <c>Microsoft.SharePoint.Client.BasePermissions</c>.
    ///
    /// PnP Core's <c>IBasePermissions</c> is deliberately not used here: it is a live data model
    /// interface whose implementation is internal and bound to a <c>PnPContext</c>, whereas the
    /// provisioning model is a plain serializable POCO layer that must round trip a template without
    /// any connection. Conversion to <c>IBasePermissions</c> happens in the object handlers, using
    /// <see cref="High"/> and <see cref="Low"/>.
    ///
    /// The behaviour below was characterised against the real CSOM assembly rather than assumed,
    /// because two parts of it are surprising and are relied upon by the schema resolvers:
    ///
    /// 1. <see cref="PermissionKind.FullMask"/> does **not** set every permission bit. CSOM sets
    ///    high = 0x7FFF and low = 0xFFFF, so for example <c>Has(PermissionKind.ManageWeb)</c> - which
    ///    is bit 30 of the low word - returns false even after <c>Set(FullMask)</c>. "Fixing" this
    ///    would change which permissions a template with Rights="FullMask" round trips to.
    /// 2. <c>Has(PermissionKind.EmptyMask)</c> always returns true, including on a fresh instance.
    ///
    /// A differential test against the CSOM implementation covers every <see cref="PermissionKind"/>
    /// value; see the provisioning test project.
    /// </remarks>
    public sealed class BasePermissions : IEquatable<BasePermissions>
    {
        private const uint FullMaskHigh = 0x7FFF;
        private const uint FullMaskLow = 0xFFFF;

        /// <summary>
        /// Gets or sets the high word of the permission mask (permissions 33 to 64).
        /// </summary>
        public uint High { get; set; }

        /// <summary>
        /// Gets or sets the low word of the permission mask (permissions 1 to 32).
        /// </summary>
        public uint Low { get; set; }

        /// <summary>
        /// Adds the given permission to this permission set.
        /// </summary>
        /// <param name="perm">The permission to add</param>
        public void Set(PermissionKind perm)
        {
            if (perm == PermissionKind.FullMask)
            {
                High = FullMaskHigh;
                Low = FullMaskLow;
                return;
            }

            // Matches CSOM: Set(EmptyMask) resets the mask rather than doing nothing. It is not the
            // inverse of Clear(EmptyMask), which IS a no-op.
            if (perm == PermissionKind.EmptyMask)
            {
                High = 0;
                Low = 0;
                return;
            }

            int index = (int)perm - 1;
            if (index < 32)
            {
                Low |= (uint)1 << index;
            }
            else
            {
                High |= (uint)1 << (index - 32);
            }
        }

        /// <summary>
        /// Removes the given permission from this permission set.
        /// </summary>
        /// <param name="perm">The permission to remove</param>
        public void Clear(PermissionKind perm)
        {
            // Matches CSOM: neither FullMask nor EmptyMask clears anything. Clear(FullMask) in
            // particular is a no-op rather than the "remove everything" it reads like - use
            // ClearAll() for that. Verified against the CSOM implementation.
            if (perm == PermissionKind.FullMask || perm == PermissionKind.EmptyMask)
            {
                return;
            }

            int index = (int)perm - 1;
            if (index < 32)
            {
                Low &= ~((uint)1 << index);
            }
            else
            {
                High &= ~((uint)1 << (index - 32));
            }
        }

        /// <summary>
        /// Removes all permissions from this permission set.
        /// </summary>
        public void ClearAll()
        {
            High = 0;
            Low = 0;
        }

        /// <summary>
        /// Determines whether this permission set contains the given permission.
        /// </summary>
        /// <param name="perm">The permission to test for</param>
        /// <returns>True if the permission is present</returns>
        public bool Has(PermissionKind perm)
        {
            // Matches CSOM: EmptyMask is always considered present.
            if (perm == PermissionKind.EmptyMask)
            {
                return true;
            }

            if (perm == PermissionKind.FullMask)
            {
                return High == FullMaskHigh && Low == FullMaskLow;
            }

            int index = (int)perm - 1;
            if (index < 32)
            {
                return (Low & ((uint)1 << index)) != 0;
            }

            return (High & ((uint)1 << (index - 32))) != 0;
        }

        /// <summary>
        /// Determines whether this permission set contains all the bits of the given mask.
        /// </summary>
        /// <param name="high">High word of the mask to test for</param>
        /// <param name="low">Low word of the mask to test for</param>
        /// <returns>True if every bit of the given mask is present</returns>
        public bool HasPermissions(uint high, uint low)
        {
            return (High & high) == high && (Low & low) == low;
        }

        /// <summary>
        /// Compares this permission set with another one.
        /// </summary>
        /// <param name="other">The permission set to compare with</param>
        /// <returns>True if both represent the same permissions</returns>
        public bool Equals(BasePermissions other)
        {
            if (other is null)
            {
                return false;
            }

            return High == other.High && Low == other.Low;
        }

        /// <summary>
        /// Compares this permission set with another object.
        /// </summary>
        /// <param name="obj">The object to compare with</param>
        /// <returns>True if the object is an equivalent permission set</returns>
        public override bool Equals(object obj)
        {
            return Equals(obj as BasePermissions);
        }

        /// <summary>
        /// Returns the hash code of this permission set.
        /// </summary>
        /// <returns>A hash code based on the permission mask</returns>
        public override int GetHashCode()
        {
            return $"{High}|{Low}".GetHashCode();
        }
    }
}
