using Newtonsoft.Json;

namespace AddUsers.Configuration
{
    /// <summary>
    /// Persisted add-in configuration (serialized to %APPDATA%\AddUsers\settings.json).
    /// </summary>
    public class AddInSettings
    {
        /// <summary>Tenant base URL, e.g. https://contoso.sharepoint.com.</summary>
        public string BaseUrl { get; set; } = string.Empty;

        /// <summary>Azure AD application (client) id used for auth.</summary>
        public string ClientId { get; set; } = string.Empty;

        /// <summary>Azure AD tenant id (GUID or contoso.onmicrosoft.com).</summary>
        public string TenantId { get; set; } = string.Empty;

        /// <summary>URL of the selected site collection.</summary>
        public string SiteUrl { get; set; } = string.Empty;

        /// <summary>Display title of the selected site collection.</summary>
        public string SiteTitle { get; set; } = string.Empty;

        /// <summary>Id of the selected SharePoint group (0 = none).</summary>
        public int GroupId { get; set; }

        /// <summary>Display name of the selected SharePoint group.</summary>
        public string GroupName { get; set; } = string.Empty;

        /// <summary>True when enough is configured to add users to a group. TenantId is optional.</summary>
        [JsonIgnore]
        public bool IsConfigured =>
            !string.IsNullOrWhiteSpace(ClientId) &&
            !string.IsNullOrWhiteSpace(SiteUrl) &&
            (GroupId > 0 || !string.IsNullOrWhiteSpace(GroupName));
    }
}
