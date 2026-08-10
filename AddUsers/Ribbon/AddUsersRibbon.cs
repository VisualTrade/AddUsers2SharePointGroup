using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using AddUsers.Configuration;
using AddUsers.Forms;
using AddUsers.Services;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AddUsers
{
    [ComVisible(true)]
    public class AddUsersRibbon : Office.IRibbonExtensibility
    {
        private const string RibbonXmlResource = "AddUsers.Ribbon.AddUsersRibbon.xml";
        private const string ImageResourceFormat = "AddUsers.Resources.{0}.png";
        private const string Caption = "Add Users to SharePoint Group";
        private const int MaxEmailsInConfirmation = 20;

        private Office.IRibbonUI _ribbon;
        private bool _busy;

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        /// <summary>Wraps the Outlook window handle so post-await dialogs stay owned by (and in front of) Outlook.</summary>
        private sealed class Win32Window : IWin32Window
        {
            private readonly IntPtr _handle;
            public Win32Window(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            // Explorer only; no customization for inspectors or other windows.
            if (!string.Equals(ribbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return GetResourceText(RibbonXmlResource);
        }

        #endregion

        #region Ribbon callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public stdole.IPictureDisp LoadImage(string imageName)
        {
            string resourceName = string.Format(ImageResourceFormat, imageName);
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                using (Image image = Image.FromStream(stream))
                {
                    return PictureConverter.ImageToPictureDisp(image);
                }
            }
        }

        public async void OnAddUsersClick(Office.IRibbonControl control)
        {
            if (_busy)
            {
                return;
            }

            _busy = true;
            try
            {
                // Safety net in case anything ran before ThisAddIn_Startup installed the
                // WinForms context: without one, the post-await code would resume on an
                // MTA thread-pool thread and the dialogs would misbehave.
                if (!(System.Threading.SynchronizationContext.Current is WindowsFormsSynchronizationContext))
                {
                    System.Threading.SynchronizationContext.SetSynchronizationContext(
                        new WindowsFormsSynchronizationContext());
                }

                // Captured while Outlook is still the foreground window; after the await the
                // browser (interactive sign-in) is usually foreground instead.
                IWin32Window owner = new Win32Window(GetActiveWindow());

                AddInSettings settings = EnsureConfigured();
                if (settings == null)
                {
                    return;
                }

                List<string> resolved;
                List<string> unresolved;
                int selectionCount;
                Outlook.MailItem mail = OutlookMailHelper.GetSelectedMailItem(
                    Globals.ThisAddIn.Application, out selectionCount);
                if (mail == null)
                {
                    MessageBox.Show(
                        "Select an email message in the mail list first.",
                        Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    resolved = OutlookMailHelper.ResolveRecipientSmtpAddresses(mail, includeSender: true, out unresolved);
                }
                finally
                {
                    Marshal.ReleaseComObject(mail);
                }

                if (resolved.Count == 0 && unresolved.Count == 0)
                {
                    MessageBox.Show(
                        "No sender or recipients could be resolved from the selected email.",
                        Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (resolved.Count == 0)
                {
                    // Nothing to send to SharePoint; still surface the failures.
                    ShowResults(owner, settings.GroupName, BuildUnresolvedOutcomes(unresolved));
                    return;
                }

                if (!ConfirmAdd(settings, resolved, unresolved, selectionCount))
                {
                    return;
                }

                List<AddUserOutcome> outcomes =
                    await SharePointService.Instance.AddUsersToGroupAsync(settings, resolved);
                if (outcomes == null)
                {
                    outcomes = new List<AddUserOutcome>();
                }

                outcomes.AddRange(BuildUnresolvedOutcomes(unresolved));
                ShowResults(owner, settings.GroupName, outcomes);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    new Win32Window(GetActiveWindow()),
                    "Adding users failed: " + ex.Message,
                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _busy = false;
            }
        }

        public void OnConfigureClick(Office.IRibbonControl control)
        {
            try
            {
                ShowConfigurationDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Configuration failed: " + ex.Message,
                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Flow helpers

        private static AddInSettings EnsureConfigured()
        {
            AddInSettings settings = SettingsStore.Load();
            if (settings != null && settings.IsConfigured)
            {
                return settings;
            }

            DialogResult choice = MessageBox.Show(
                "The add-in is not configured yet. Configure it now?",
                Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (choice != DialogResult.Yes)
            {
                return null;
            }

            ShowConfigurationDialog();
            settings = SettingsStore.Load();
            return settings != null && settings.IsConfigured ? settings : null;
        }

        private static bool ConfirmAdd(AddInSettings settings, List<string> resolved, List<string> unresolved, int selectionCount)
        {
            string site = string.IsNullOrWhiteSpace(settings.SiteTitle) ? settings.SiteUrl : settings.SiteTitle;
            var message = new StringBuilder();
            message.AppendFormat(
                "Add {0} user(s) to SharePoint group '{1}' on {2}?", resolved.Count, settings.GroupName, site);
            message.AppendLine();
            if (selectionCount > 1)
            {
                message.AppendFormat(
                    "({0} messages are selected; only the first one's recipients are used.)", selectionCount);
                message.AppendLine();
            }
            message.AppendLine();

            int shown = Math.Min(resolved.Count, MaxEmailsInConfirmation);
            for (int i = 0; i < shown; i++)
            {
                message.AppendLine(resolved[i]);
            }

            if (resolved.Count > shown)
            {
                message.AppendFormat("... and {0} more", resolved.Count - shown);
                message.AppendLine();
            }

            if (unresolved.Count > 0)
            {
                message.AppendLine();
                message.AppendFormat(
                    "{0} recipient(s) could not be resolved to an SMTP address and will be skipped.",
                    unresolved.Count);
                message.AppendLine();
            }

            return MessageBox.Show(
                message.ToString(), Caption,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private static List<AddUserOutcome> BuildUnresolvedOutcomes(List<string> unresolved)
        {
            var outcomes = new List<AddUserOutcome>(unresolved.Count);
            foreach (string name in unresolved)
            {
                outcomes.Add(new AddUserOutcome
                {
                    Email = name,
                    Success = false,
                    Detail = "Could not be resolved to an SMTP address."
                });
            }

            return outcomes;
        }

        private static void ShowResults(IWin32Window owner, string groupName, List<AddUserOutcome> outcomes)
        {
            using (var form = new ResultsForm(groupName, outcomes))
            {
                form.ShowDialog(owner);
            }
        }

        private static void ShowConfigurationDialog()
        {
            using (var form = new ConfigurationForm())
            {
                form.ShowDialog();
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        #endregion
    }
}
