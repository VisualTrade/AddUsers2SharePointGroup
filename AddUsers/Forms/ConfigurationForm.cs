using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AddUsers.Configuration;
using AddUsers.Services;

namespace AddUsers.Forms
{
    /// <summary>
    /// Modal dialog for configuring the SharePoint connection and the target site/group.
    /// Returns DialogResult.OK after a successful save.
    /// </summary>
    public class ConfigurationForm : Form
    {
        private readonly AddInSettings _settings;

        private TextBox _baseUrlBox;
        private TextBox _clientIdBox;
        private TextBox _tenantIdBox;
        private Button _loadSitesButton;
        private ComboBox _siteCombo;
        private ComboBox _groupCombo;
        private Label _currentLabel;
        private Button _saveButton;
        private Button _cancelButton;

        private int _row;
        private bool _busy;
        private bool _suppressEvents;

        public ConfigurationForm()
        {
            _settings = SettingsStore.Load();
            BuildLayout();
            Prefill();
        }

        /// <summary>The settings instance edited by this dialog.</summary>
        public AddInSettings Settings => _settings;

        private void BuildLayout()
        {
            Text = "Add Users - Configuration";
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 312);

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                Padding = new Padding(12, 12, 12, 0)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            _baseUrlBox = MakeTextBox();
            _clientIdBox = MakeTextBox();
            _tenantIdBox = MakeTextBox();

            _loadSitesButton = new Button
            {
                Text = "Load sites",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 3, 0, 9)
            };
            _loadSitesButton.Click += OnLoadSitesClick;

            _siteCombo = MakeCombo();
            _siteCombo.SelectedIndexChanged += OnSiteSelectionChanged;
            _siteCombo.Enabled = false;

            _groupCombo = MakeCombo();
            _groupCombo.Enabled = false;

            _currentLabel = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 9, 0, 0)
            };

            AddRow(table, "Site / base URL:", _baseUrlBox);
            AddRow(table, "Client ID:", _clientIdBox);
            AddRow(table, "Tenant ID (optional):", _tenantIdBox);
            AddRow(table, null, _loadSitesButton);
            AddRow(table, "Site:", _siteCombo);
            AddRow(table, "Group:", _groupCombo);

            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(_currentLabel, 0, _row);
            table.SetColumnSpan(_currentLabel, 2);

            _saveButton = new Button
            {
                Text = "Save",
                Size = new Size(88, 27),
                Margin = new Padding(6, 0, 0, 0)
            };
            _saveButton.Click += OnSaveClick;

            _cancelButton = new Button
            {
                Text = "Cancel",
                Size = new Size(88, 27),
                Margin = new Padding(6, 0, 0, 0),
                DialogResult = DialogResult.Cancel
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12, 8, 12, 12)
            };
            buttons.Controls.Add(_cancelButton);
            buttons.Controls.Add(_saveButton);

            Controls.Add(table);
            Controls.Add(buttons);
            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
        }

        private static TextBox MakeTextBox()
        {
            return new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 0, 3)
            };
        }

        private static ComboBox MakeCombo()
        {
            return new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 0, 3),
                DisplayMember = "Title"
            };
        }

        private void AddRow(TableLayoutPanel table, string labelText, Control control)
        {
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (labelText != null)
            {
                table.Controls.Add(new Label
                {
                    Text = labelText,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 7, 6, 0)
                }, 0, _row);
            }
            table.Controls.Add(control, 1, _row);
            _row++;
        }

        private void Prefill()
        {
            _baseUrlBox.Text = _settings.BaseUrl ?? string.Empty;
            _clientIdBox.Text = _settings.ClientId ?? string.Empty;
            _tenantIdBox.Text = _settings.TenantId ?? string.Empty;

            string site = string.IsNullOrWhiteSpace(_settings.SiteTitle) ? _settings.SiteUrl : _settings.SiteTitle;
            _currentLabel.Text =
                string.IsNullOrWhiteSpace(site) || string.IsNullOrWhiteSpace(_settings.GroupName)
                    ? "Currently: not configured"
                    : "Currently: " + site + " / " + _settings.GroupName;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            UseWaitCursor = busy;
            _baseUrlBox.Enabled = !busy;
            _clientIdBox.Enabled = !busy;
            _tenantIdBox.Enabled = !busy;
            _loadSitesButton.Enabled = !busy;
            _siteCombo.Enabled = !busy && _siteCombo.Items.Count > 0;
            _groupCombo.Enabled = !busy && _groupCombo.Items.Count > 0;
            _saveButton.Enabled = !busy;
            // Cancel stays enabled so the dialog can be closed even if an interactive
            // sign-in is abandoned in the browser and the operation never completes.
        }

        private async void OnLoadSitesClick(object sender, EventArgs e)
        {
            if (_busy)
                return;

            string baseUrl = _baseUrlBox.Text.Trim();
            string clientId = _clientIdBox.Text.Trim();
            string tenantId = _tenantIdBox.Text.Trim();
            if (baseUrl.Length == 0 || clientId.Length == 0)
            {
                MessageBox.Show(this, "Enter the base URL and client id first.",
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetBusy(true);
            List<SiteInfo> sites = null;
            string discoveryError = null;
            try
            {
                // No ResetAuth here: the service caches AuthenticationManagers per
                // (ClientId, TenantId), so changed credentials get a fresh sign-in
                // automatically while repeat clicks reuse the cached MSAL token.
                try
                {
                    var loaded = await SharePointService.Instance.GetSitesAsync(baseUrl, clientId, tenantId);
                    sites = loaded.ToList();
                }
                catch (Exception ex)
                {
                    discoveryError = ex.Message;
                }

                // Tenant-wide discovery fails (or finds nothing) when the app grant is
                // Sites.Selected — only explicitly granted sites are reachable. Fall
                // back to treating the entered URL as the site itself.
                if (sites == null || sites.Count == 0)
                {
                    try
                    {
                        var single = await SharePointService.Instance.GetSiteAsync(baseUrl, clientId, tenantId);
                        sites = new List<SiteInfo> { single };
                    }
                    catch (Exception ex)
                    {
                        if (!IsDisposed)
                        {
                            var message = new System.Text.StringBuilder("Could not load sites.");
                            if (discoveryError != null)
                                message.Append("\r\n\r\nSite discovery: ").Append(discoveryError);
                            message.Append("\r\n\r\nDirect site load: ").Append(ex.Message);
                            message.Append("\r\n\r\nIf the app registration uses Sites.Selected, enter the FULL URL of a granted site (e.g. https://tenant.sharepoint.com/sites/yoursite).");
                            MessageBox.Show(this, message.ToString(),
                                "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        sites = null;
                    }
                }
            }
            finally
            {
                if (!IsDisposed)
                    SetBusy(false);
            }

            if (IsDisposed || sites == null)
                return;

            BindSites(sites);
        }

        private void BindSites(List<SiteInfo> sites)
        {
            _suppressEvents = true;
            try
            {
                _siteCombo.BeginUpdate();
                _siteCombo.Items.Clear();
                foreach (var site in sites)
                    _siteCombo.Items.Add(site);
                _siteCombo.EndUpdate();
                _siteCombo.SelectedIndex = -1;

                _groupCombo.Items.Clear();
                _groupCombo.Enabled = false;
            }
            finally
            {
                _suppressEvents = false;
            }

            _siteCombo.Enabled = sites.Count > 0;
            if (sites.Count == 0)
            {
                MessageBox.Show(this,
                    "No sites were found. Site discovery uses SharePoint search, so newly " +
                    "created sites can take a while to appear; also check the base URL.",
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Re-select the previously saved site, which triggers the groups load.
            if (!string.IsNullOrEmpty(_settings.SiteUrl))
            {
                for (int i = 0; i < sites.Count; i++)
                {
                    if (string.Equals(sites[i].Url, _settings.SiteUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        _siteCombo.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private async void OnSiteSelectionChanged(object sender, EventArgs e)
        {
            if (_suppressEvents || _busy)
                return;

            var site = _siteCombo.SelectedItem as SiteInfo;
            if (site == null)
                return;

            SetBusy(true);

            // Clear the previous site's groups up front so a failed fetch can never leave
            // stale, selectable entries behind (which Save would pair with the new site).
            _suppressEvents = true;
            try
            {
                _groupCombo.Items.Clear();
                _groupCombo.SelectedIndex = -1;
            }
            finally
            {
                _suppressEvents = false;
            }
            _groupCombo.Enabled = false;

            List<GroupInfo> groups = null;
            try
            {
                var loaded = await SharePointService.Instance.GetGroupsAsync(
                    site.Url, _clientIdBox.Text.Trim(), _tenantIdBox.Text.Trim());
                groups = loaded.ToList();
            }
            catch (Exception ex)
            {
                if (!IsDisposed)
                    MessageBox.Show(this, "Could not load groups:\r\n" + ex.Message,
                        "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!IsDisposed)
                    SetBusy(false);
            }

            if (IsDisposed || groups == null)
                return;

            _suppressEvents = true;
            try
            {
                _groupCombo.BeginUpdate();
                _groupCombo.Items.Clear();
                foreach (var group in groups)
                    _groupCombo.Items.Add(group);
                _groupCombo.EndUpdate();
                _groupCombo.SelectedIndex = -1;
            }
            finally
            {
                _suppressEvents = false;
            }

            _groupCombo.Enabled = groups.Count > 0;

            // Only re-select the saved group when this is the saved site: group ids are
            // site-relative small integers, so the same id on another site is an
            // unrelated group.
            if (_settings.GroupId > 0 &&
                string.Equals(site.Url, _settings.SiteUrl, StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < groups.Count; i++)
                {
                    if (groups[i].Id == _settings.GroupId)
                    {
                        _groupCombo.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void OnSaveClick(object sender, EventArgs e)
        {
            if (_busy)
                return;

            string baseUrl = _baseUrlBox.Text.Trim();
            string clientId = _clientIdBox.Text.Trim();
            string tenantId = _tenantIdBox.Text.Trim();
            if (baseUrl.Length == 0 || clientId.Length == 0)
            {
                MessageBox.Show(this, "Enter the base URL and client id.",
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var site = _siteCombo.SelectedItem as SiteInfo;
            var group = _groupCombo.SelectedItem as GroupInfo;

            if (site != null && group == null)
            {
                MessageBox.Show(this, "Choose a group for the selected site.",
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // A saved site/group may only be kept when it still lives under the entered
            // base URL; otherwise a changed base URL would silently retain the old
            // tenant's site and group.
            bool hasSaved = !string.IsNullOrEmpty(_settings.SiteUrl) && _settings.GroupId > 0;
            Uri newBase, savedSite;
            bool savedSiteMatchesBase = hasSaved
                && Uri.TryCreate(baseUrl, UriKind.Absolute, out newBase)
                && Uri.TryCreate(_settings.SiteUrl, UriKind.Absolute, out savedSite)
                && string.Equals(newBase.Host, savedSite.Host, StringComparison.OrdinalIgnoreCase);
            if (site == null && !savedSiteMatchesBase)
            {
                MessageBox.Show(this,
                    hasSaved
                        ? "The base URL changed. Click 'Load sites' and choose a site and group before saving."
                        : "Load the sites and choose a site and group before saving.",
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _settings.BaseUrl = baseUrl;
            _settings.ClientId = clientId;
            _settings.TenantId = tenantId;
            if (site != null && group != null)
            {
                _settings.SiteUrl = site.Url;
                _settings.SiteTitle = site.Title;
                _settings.GroupId = group.Id;
                _settings.GroupName = group.Title;
            }

            try
            {
                SettingsStore.Save(_settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not save the settings:\r\n" + ex.Message,
                    "Add Users", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
