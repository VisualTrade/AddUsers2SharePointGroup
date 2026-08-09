using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AddUsers.Services;

namespace AddUsers.Forms
{
    /// <summary>
    /// Modal dialog showing the outcome of adding recipients to a SharePoint group:
    /// a summary line plus one row per e-mail address, with failures highlighted.
    /// </summary>
    public class ResultsForm : Form
    {
        private ListView _listView;

        public ResultsForm(string groupName, IEnumerable<AddUserOutcome> results)
        {
            List<AddUserOutcome> items = (results ?? Enumerable.Empty<AddUserOutcome>())
                .Where(r => r != null)
                .ToList();

            int succeeded = items.Count(r => r.Success);
            int failed = items.Count - succeeded;

            BuildLayout(groupName, items.Count, succeeded, failed);
            PopulateList(items);
        }

        private void BuildLayout(string groupName, int total, int succeeded, int failed)
        {
            Text = "Add Users - Results";
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(600, 400);
            MinimumSize = new Size(480, 320);

            string target = string.IsNullOrWhiteSpace(groupName) ? "the group" : "'" + groupName + "'";
            string text = succeeded + " of " + total + " user(s) added to " + target + ".";
            if (failed > 0)
                text += " " + failed + " failed.";

            var summary = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 36,
                Padding = new Padding(12, 12, 12, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = failed > 0 ? Color.Firebrick : SystemColors.ControlText,
                Text = text
            };

            _listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                Margin = new Padding(0)
            };
            _listView.Columns.Add("Email", 220);
            _listView.Columns.Add("Status", 90);
            _listView.Columns.Add("Detail", 240);
            _listView.Resize += (s, e) => FitDetailColumn();
            // Refit once the handle exists: at construction ClientSize does not yet account
            // for the vertical scrollbar, which would otherwise cause a spurious horizontal one.
            Shown += (s, e) => FitDetailColumn();

            var listHost = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 8, 12, 0)
            };
            listHost.Controls.Add(_listView);

            var okButton = new Button
            {
                Text = "OK",
                Size = new Size(88, 27),
                DialogResult = DialogResult.OK
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12, 8, 12, 12)
            };
            buttons.Controls.Add(okButton);

            Controls.Add(listHost);
            Controls.Add(summary);
            Controls.Add(buttons);
            AcceptButton = okButton;
            CancelButton = okButton;
        }

        private void PopulateList(List<AddUserOutcome> items)
        {
            _listView.BeginUpdate();
            foreach (var result in items)
            {
                var item = new ListViewItem(result.Email ?? string.Empty);
                item.SubItems.Add(result.Success ? "Added" : "Failed");
                item.SubItems.Add(result.Detail ?? string.Empty);
                if (!result.Success)
                {
                    item.BackColor = Color.MistyRose;
                    item.ForeColor = Color.Firebrick;
                }
                _listView.Items.Add(item);
            }
            _listView.EndUpdate();
            FitDetailColumn();
        }

        private void FitDetailColumn()
        {
            if (_listView == null || _listView.Columns.Count < 3)
                return;

            int remaining = _listView.ClientSize.Width
                - _listView.Columns[0].Width
                - _listView.Columns[1].Width;
            if (remaining > 80)
                _listView.Columns[2].Width = remaining;
        }
    }
}
