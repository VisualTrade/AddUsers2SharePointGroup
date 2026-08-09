using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AddUsers.Configuration;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using PnP.Framework;

namespace AddUsers.Services
{
    /// <summary>
    /// Singleton facade over PnP.Framework / CSOM for the operations the add-in needs:
    /// discovering sites, listing site groups and adding users to a group.
    /// </summary>
    public sealed class SharePointService
    {
        private const string DefaultTenant = "organizations";
        private const string RedirectUrl = "http://localhost";
        private const string RelevantResultsTableType = "RelevantResults";
        private const int SearchPageSize = 500;
        private const int SearchMaxRows = 2000;

        private static readonly SharePointService _instance = new SharePointService();

        public static SharePointService Instance
        {
            get { return _instance; }
        }

        // One AuthenticationManager per (ClientId, TenantId) so MSAL's token cache is reused
        // and the user is only prompted interactively when no valid token exists.
        private readonly ConcurrentDictionary<string, AuthenticationManager> _authManagers =
            new ConcurrentDictionary<string, AuthenticationManager>(StringComparer.OrdinalIgnoreCase);

        private SharePointService()
        {
        }

        /// <summary>
        /// Drops all cached AuthenticationManagers so the next call signs in from scratch.
        /// Call when the configured Client ID or Tenant ID changes.
        /// </summary>
        public void ResetAuth()
        {
            foreach (string key in _authManagers.Keys.ToList())
            {
                AuthenticationManager manager;
                if (_authManagers.TryRemove(key, out manager))
                {
                    try { manager.Dispose(); }
                    catch { /* best effort */ }
                }
            }
        }

        /// <summary>
        /// Discovers sites under <paramref name="baseSiteUrl"/> using SharePoint search
        /// (contentclass:STS_Site), deduplicated by URL and sorted by title.
        /// </summary>
        public async Task<List<SiteInfo>> GetSitesAsync(string baseSiteUrl, string clientId, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(baseSiteUrl))
                throw new ArgumentException("A base site URL is required.", nameof(baseSiteUrl));

            string baseUrl = baseSiteUrl.Trim().TrimEnd('/');

            try
            {
                using (ClientContext ctx = await GetContextAsync(clientId, tenantId, baseUrl).ConfigureAwait(false))
                {
                    var sitesByUrl = new Dictionary<string, SiteInfo>(StringComparer.OrdinalIgnoreCase);

                    // Page through the results: a single query is capped at 500 rows by
                    // SharePoint Online, which would silently hide sites on large tenants.
                    for (int startRow = 0; startRow < SearchMaxRows; startRow += SearchPageSize)
                    {
                        var query = new KeywordQuery(ctx)
                        {
                            QueryText = "contentclass:STS_Site path:\"" + baseUrl + "\"",
                            RowLimit = SearchPageSize,
                            StartRow = startRow,
                            TrimDuplicates = false
                        };
                        query.SelectProperties.Add("Title");
                        query.SelectProperties.Add("Path");
                        query.SelectProperties.Add("SPSiteUrl");

                        var executor = new SearchExecutor(ctx);
                        ClientResult<ResultTableCollection> results = executor.ExecuteQuery(query);
                        await ctx.ExecuteQueryRetryAsync().ConfigureAwait(false);

                        ResultTable table =
                            results.Value.FirstOrDefault(t =>
                                string.Equals(t.TableType, RelevantResultsTableType, StringComparison.OrdinalIgnoreCase))
                            ?? results.Value.FirstOrDefault();

                        int rowCount = 0;
                        if (table != null)
                        {
                            foreach (IDictionary<string, object> row in table.ResultRows)
                            {
                                rowCount++;

                                string url = GetRowValue(row, "SPSiteUrl");
                                if (string.IsNullOrEmpty(url))
                                    url = GetRowValue(row, "Path");
                                if (string.IsNullOrEmpty(url))
                                    continue;

                                url = url.TrimEnd('/');
                                if (sitesByUrl.ContainsKey(url))
                                    continue;

                                string title = GetRowValue(row, "Title");
                                sitesByUrl[url] = new SiteInfo
                                {
                                    Title = string.IsNullOrEmpty(title) ? url : title,
                                    Url = url
                                };
                            }
                        }

                        if (rowCount < SearchPageSize)
                            break;
                    }

                    return sitesByUrl.Values
                        .OrderBy(s => s.Title, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(s => s.Url, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
            catch (Exception ex) when (IsAuthFailure(ex))
            {
                throw CreateAuthError(ex);
            }
        }

        /// <summary>Lists the site groups of the given site (Id and Title only), sorted by title.</summary>
        public async Task<List<GroupInfo>> GetGroupsAsync(string siteUrl, string clientId, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
                throw new ArgumentException("A site URL is required.", nameof(siteUrl));

            try
            {
                using (ClientContext ctx = await GetContextAsync(clientId, tenantId, siteUrl.Trim()).ConfigureAwait(false))
                {
                    GroupCollection groups = ctx.Web.SiteGroups;
                    ctx.Load(groups, gs => gs.Include(g => g.Id, g => g.Title));
                    await ctx.ExecuteQueryRetryAsync().ConfigureAwait(false);

                    return groups
                        .Select(g => new GroupInfo { Id = g.Id, Title = g.Title })
                        .OrderBy(g => g.Title, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
            catch (Exception ex) when (IsAuthFailure(ex))
            {
                throw CreateAuthError(ex);
            }
        }

        /// <summary>
        /// Adds each email address to the group configured in <paramref name="settings"/>.
        /// The group is resolved by GroupId when greater than zero, otherwise by GroupName.
        /// Failures are reported per address; one bad address does not stop the rest.
        /// </summary>
        public async Task<List<AddUserOutcome>> AddUsersToGroupAsync(AddInSettings settings, List<string> emails)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            string clientId = settings.ClientId;
            string tenantId = settings.TenantId;
            string siteUrl = settings.SiteUrl;
            int groupId = settings.GroupId;
            string groupName = settings.GroupName;

            if (string.IsNullOrWhiteSpace(siteUrl))
                throw new ArgumentException("A site URL is required.", nameof(settings));
            if (groupId <= 0 && string.IsNullOrWhiteSpace(groupName))
                throw new ArgumentException("Either a group ID or a group name is required.", nameof(settings));

            var results = new List<AddUserOutcome>();

            try
            {
                using (ClientContext ctx = await GetContextAsync(clientId, tenantId, siteUrl.Trim()).ConfigureAwait(false))
                {
                    Group group = groupId > 0
                        ? ctx.Web.SiteGroups.GetById(groupId)
                        : ctx.Web.SiteGroups.GetByName(groupName.Trim());

                    try
                    {
                        ctx.Load(group, g => g.Id, g => g.Title);
                        await ctx.ExecuteQueryRetryAsync().ConfigureAwait(false);
                    }
                    catch (ServerException ex)
                    {
                        string label = groupId > 0
                            ? "with ID " + groupId
                            : "named '" + groupName.Trim() + "'";
                        throw new InvalidOperationException(
                            "The SharePoint group " + label + " could not be found on " + siteUrl.Trim() +
                            ". (" + ex.Message + ")", ex);
                    }

                    List<string> pending = (emails ?? new List<string>())
                        .Select(e => (e ?? string.Empty).Trim())
                        .Where(e => e.Length > 0)
                        .ToList();

                    for (int i = 0; i < pending.Count; i++)
                    {
                        string email = pending[i];
                        try
                        {
                            // A bare address lets SharePoint Online resolve by UPN or mail
                            // attribute, which also covers guests (#EXT# accounts) and users
                            // whose primary SMTP differs from their UPN; a hard-coded
                            // i:0#.f|membership| claim would break both.
                            User user = ctx.Web.EnsureUser(email);
                            group.Users.AddUser(user);
                            await ctx.ExecuteQueryRetryAsync().ConfigureAwait(false);

                            results.Add(new AddUserOutcome { Email = email, Success = true, Detail = "Added" });
                        }
                        catch (Exception ex) when (!IsAuthFailure(ex))
                        {
                            results.Add(new AddUserOutcome
                            {
                                Email = email,
                                Success = false,
                                Detail = ex.Message
                            });
                        }
                        catch (Exception ex) when (IsAuthFailure(ex))
                        {
                            // Sign-in died mid-batch. Users added so far are committed on the
                            // server; preserve their outcomes instead of throwing them away.
                            if (results.Count == 0)
                                throw;

                            results.Add(new AddUserOutcome
                            {
                                Email = email,
                                Success = false,
                                Detail = CreateAuthError(ex).Message
                            });
                            for (int j = i + 1; j < pending.Count; j++)
                            {
                                results.Add(new AddUserOutcome
                                {
                                    Email = pending[j],
                                    Success = false,
                                    Detail = "Not attempted (sign-in failed)."
                                });
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex) when (IsAuthFailure(ex))
            {
                throw CreateAuthError(ex);
            }

            return results;
        }

        private async Task<ClientContext> GetContextAsync(string clientId, string tenantId, string siteUrl)
        {
            AuthenticationManager authManager = GetAuthenticationManager(clientId, tenantId);
            return await authManager.GetContextAsync(siteUrl).ConfigureAwait(false);
        }

        private AuthenticationManager GetAuthenticationManager(string clientId, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(clientId))
            {
                throw new InvalidOperationException(
                    "The Azure AD application (client) ID is not configured. " +
                    "Open the add-in configuration and enter a Client ID before connecting to SharePoint.");
            }

            string id = clientId.Trim();
            string tenant = string.IsNullOrWhiteSpace(tenantId) ? DefaultTenant : tenantId.Trim();
            string cacheKey = id + "|" + tenant;

            return _authManagers.GetOrAdd(cacheKey, _ =>
                AuthenticationManager.CreateWithInteractiveLogin(id, RedirectUrl, tenant));
        }

        private static string GetRowValue(IDictionary<string, object> row, string key)
        {
            object value;
            return row.TryGetValue(key, out value) && value != null ? value.ToString() : null;
        }

        private static bool IsAuthFailure(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                if (e is MsalException)
                    return true;
            }
            return false;
        }

        private static InvalidOperationException CreateAuthError(Exception ex)
        {
            return new InvalidOperationException(
                "Could not sign in to SharePoint. If you cancelled or closed the sign-in window, just " +
                "try again. Otherwise verify the Client ID and Tenant ID in the add-in configuration " +
                "and make sure the Azure AD app registration allows public client flows with the " +
                "'http://localhost' redirect URI. (" + ex.Message + ")", ex);
        }
    }
}
