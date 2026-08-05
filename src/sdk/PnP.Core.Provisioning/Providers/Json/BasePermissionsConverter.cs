using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
// PnP.Core.Model.SharePoint also declares a BasePermissions (an internal, context bound data
// model type), so the provisioning POCO is disambiguated explicitly.
using BasePermissions = PnP.Core.Provisioning.Model.BasePermissions;

namespace PnP.Core.Provisioning.Providers.Json
{
    /// <summary>
    /// Serializes a <see cref="BasePermissions"/> as the comma separated list of permission names it
    /// contains, and reads it back.
    /// </summary>
    /// <remarks>
    /// Converted from a Newtonsoft <c>JsonConverter</c> to a <c>System.Text.Json</c>
    /// <see cref="JsonConverter{T}"/>, since PnP Core has no Newtonsoft dependency.
    ///
    /// One behavioural fix was made during the conversion. The PnP Framework version read a comma
    /// separated value like this:
    ///
    ///     foreach (var pk in basePermissionString.Split(new char[] { ',' }))
    ///     {
    ///         if (Enum.TryParse(basePermissionString, out permissionKind))   // note: not pk
    ///
    /// It parsed the whole string on every iteration instead of the current item, so a value such as
    /// "ViewListItems,AddListItems" never parsed and no permissions were set. The equivalent XML
    /// resolver (FromStringToBasePermissionsValueResolver) parses the item, so the two formats
    /// disagreed. This converter parses the item.
    /// </remarks>
    internal class BasePermissionsConverter : JsonConverter<BasePermissions>
    {
        public override BasePermissions Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            BasePermissions result = new BasePermissions();

            string basePermissionString = reader.TokenType == JsonTokenType.Null ? null : reader.GetString();

            if (!string.IsNullOrEmpty(basePermissionString))
            {
                // Is it an int value (for backwards compatibility)?
                if (int.TryParse(basePermissionString, out int permissionInt))
                {
                    result.Set((PermissionKind)permissionInt);
                }
                else
                {
                    foreach (string pk in basePermissionString.Split(new char[] { ',' }))
                    {
                        if (Enum.TryParse(pk.Trim(), out PermissionKind permissionKind))
                        {
                            result.Set(permissionKind);
                        }
                    }
                }
            }

            return result;
        }

        public override void Write(Utf8JsonWriter writer, BasePermissions value, JsonSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteNullValue();
                return;
            }

            List<string> permissions = new List<string>();
            foreach (PermissionKind pk in (PermissionKind[])Enum.GetValues(typeof(PermissionKind)))
            {
                if (value.Has(pk) && pk != PermissionKind.EmptyMask)
                {
                    permissions.Add(pk.ToString());
                }
            }

            writer.WriteStringValue(string.Join(",", permissions.ToArray()));
        }
    }
}
