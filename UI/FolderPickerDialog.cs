using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>
    /// Ordnerauswahl im App-Design (der System-Dialog lässt sich nicht einfärben).
    /// Beliebige Pfade inklusive Netzwerkfreigaben können direkt eingetippt werden.
    /// </summary>
    public sealed class FolderPickerDialog : FuturisticForm {
        private const string PlaceholderKey = "\0loading";

        private readonly TreeView tree;
        private readonly ThemedTextBox pathBox;
        private readonly ImageList treeIcons;
        private readonly string initialPath;

        public FolderPickerDialog(string title, string initialPath) {
            this.initialPath = initialPath;
            Text = title;
            CaptionSubtitle = "Ordner auswählen";
            EnableMinimizeButton = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(560, 520);

            var heading = new Label {
                AutoSize = false,
                AutoEllipsis = true,
                UseMnemonic = false,
                Text = title,
                Font = new Font("Segoe UI Semibold", 13f),
                Location = new Point(24, 16),
                Size = new Size(512, 28)
            };

            var card = new CardPanel {
                Location = new Point(24, 52),
                Size = new Size(512, 346)
            };

            treeIcons = new ImageList { ImageSize = new Size(18, 18), ColorDepth = ColorDepth.Depth32Bit };
            treeIcons.Images.Add(CreateFolderIcon());

            tree = new TreeView {
                Location = new Point(10, 10),
                Size = new Size(492, 326),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.None,
                ShowLines = false,
                FullRowSelect = true,
                HideSelection = false,
                HotTracking = true,
                ItemHeight = 26,
                ImageList = treeIcons,
                PathSeparator = "\\"
            };
            tree.BeforeExpand += Tree_BeforeExpand;
            tree.AfterSelect += Tree_AfterSelect;
            card.Controls.Add(tree);

            var pathLabel = new Label {
                AutoSize = false,
                Text = "Pfad:",
                Tag = "subtle",
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(24, 411),
                Size = new Size(44, 23)
            };

            pathBox = new ThemedTextBox {
                Location = new Point(70, 411),
                Size = new Size(466, 23),
                Text = initialPath ?? string.Empty
            };

            var okButton = new ThemedButton {
                Text = "Auswählen",
                Kind = ButtonKind.Primary,
                Location = new Point(308, 468),
                Size = new Size(110, 32)
            };
            okButton.Click += OkButton_Click;

            var cancelButton = new ThemedButton {
                Text = "Abbrechen",
                DialogResult = DialogResult.Cancel,
                Location = new Point(426, 468),
                Size = new Size(110, 32)
            };

            Controls.Add(heading);
            Controls.Add(card);
            Controls.Add(pathLabel);
            Controls.Add(pathBox);
            Controls.Add(okButton);
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        /// <summary>Der gewählte Ordnerpfad nach Bestätigung mit OK.</summary>
        public string SelectedPath { get; private set; }

        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            AppTheme.ApplyNativeTheme(tree);
            LoadRoots();
            ExpandTo(initialPath);
        }

        private void OkButton_Click(object sender, EventArgs e) {
            string path = (pathBox.Text ?? string.Empty).Trim();
            if (path.Length == 0) {
                AppDialog.Show(this, "Bitte einen Ordner auswählen oder einen Pfad eingeben.", "Ordner auswählen",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            SelectedPath = path;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void LoadRoots() {
            tree.BeginUpdate();
            tree.Nodes.Clear();

            AddRoot("Desktop", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            AddRoot("Dokumente", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

            foreach (DriveInfo drive in DriveInfo.GetDrives()) {
                try {
                    if (!drive.IsReady) {
                        continue;
                    }
                    string label = string.IsNullOrWhiteSpace(drive.VolumeLabel)
                        ? drive.Name.TrimEnd('\\')
                        : $"{drive.VolumeLabel} ({drive.Name.TrimEnd('\\')})";
                    AddRoot(label, drive.RootDirectory.FullName);
                } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                    // Nicht lesbare Laufwerke werden nicht angeboten.
                }
            }

            tree.EndUpdate();
        }

        private void AddRoot(string label, string path) {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) {
                return;
            }
            tree.Nodes.Add(CreateNode(label, path));
        }

        private static TreeNode CreateNode(string label, string path) {
            var node = new TreeNode(label) { Tag = path };
            node.Nodes.Add(new TreeNode { Name = PlaceholderKey });
            return node;
        }

        private void Tree_BeforeExpand(object sender, TreeViewCancelEventArgs e) {
            LoadChildren(e.Node);
        }

        private void LoadChildren(TreeNode node) {
            if (node.Nodes.Count != 1 || node.Nodes[0].Name != PlaceholderKey) {
                return;
            }

            tree.BeginUpdate();
            node.Nodes.Clear();
            try {
                var directories = new DirectoryInfo((string)node.Tag).EnumerateDirectories()
                    .Where(d => (d.Attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0)
                    .OrderBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase);
                foreach (DirectoryInfo directory in directories) {
                    node.Nodes.Add(CreateNode(directory.Name, directory.FullName));
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException) {
                // Nicht lesbare Ordner bleiben einfach ohne Unterordner.
            }
            tree.EndUpdate();
        }

        private void Tree_AfterSelect(object sender, TreeViewEventArgs e) {
            if (e.Node?.Tag is string path) {
                pathBox.Text = path;
            }
        }

        // Klappt den Baum bis zum vorgegebenen Pfad auf (soweit lesbar).
        private void ExpandTo(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return;
            }

            string full;
            try {
                full = Path.GetFullPath(path).TrimEnd('\\');
            } catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException || ex is System.Security.SecurityException) {
                return;
            }

            string target = full;
            TreeNode current = tree.Nodes.Cast<TreeNode>()
                .Where(n => IsSameOrParent((string)n.Tag, target))
                .OrderByDescending(n => ((string)n.Tag).Length)
                .FirstOrDefault();

            while (current != null) {
                tree.SelectedNode = current;
                if (string.Equals(((string)current.Tag).TrimEnd('\\'), target, StringComparison.OrdinalIgnoreCase)) {
                    break;
                }
                LoadChildren(current);
                current.Expand();
                current = current.Nodes.Cast<TreeNode>()
                    .FirstOrDefault(n => n.Tag is string childPath && IsSameOrParent(childPath, target));
            }

            pathBox.Text = path;
            tree.SelectedNode?.EnsureVisible();
        }

        private static bool IsSameOrParent(string candidate, string path) {
            string normalized = candidate.TrimEnd('\\');
            return string.Equals(normalized, path, StringComparison.OrdinalIgnoreCase)
                || path.StartsWith(normalized + "\\", StringComparison.OrdinalIgnoreCase);
        }

        private static Bitmap CreateFolderIcon() {
            var bitmap = new Bitmap(18, 18);
            using (Graphics g = Graphics.FromImage(bitmap)) {
                Icons.Draw(g, IconKind.Folder, new RectangleF(1f, 1f, 16f, 16f), AppTheme.Accent, 1.4f);
            }
            return bitmap;
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                treeIcons.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
