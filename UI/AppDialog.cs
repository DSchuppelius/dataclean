using System;
using System.Collections.Generic;
using System.Drawing;
using System.Media;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>
    /// Ersatz für <see cref="MessageBox"/> im App-Design. Gleiche Parameter und
    /// Rückgabewerte, damit Aufrufer unverändert auf das DialogResult prüfen können.
    /// </summary>
    public static class AppDialog {
        public static DialogResult Show(IWin32Window owner, string text, string caption,
            MessageBoxButtons buttons = MessageBoxButtons.OK,
            MessageBoxIcon icon = MessageBoxIcon.None,
            MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1) {
            using (var dialog = new MessageDialogForm(text, caption, buttons, icon, defaultButton)) {
                Form ownerForm = owner as Form ?? Form.ActiveForm;
                dialog.TopMost = ownerForm != null && ownerForm.TopMost;
                dialog.StartPosition = ownerForm != null && ownerForm.Visible
                    ? FormStartPosition.CenterParent
                    : FormStartPosition.CenterScreen;
                return dialog.ShowDialog(owner);
            }
        }

        public static DialogResult Show(string text, string caption,
            MessageBoxButtons buttons = MessageBoxButtons.OK,
            MessageBoxIcon icon = MessageBoxIcon.None,
            MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1) {
            return Show(null, text, caption, buttons, icon, defaultButton);
        }
    }

    internal sealed class MessageDialogForm : FuturisticForm {
        private const int DialogWidth = 500;
        private const int Gutter = 24;
        private const int IconBoxSize = 40;
        private const int TextLeft = Gutter + IconBoxSize + 16;
        private const int ButtonHeight = 32;

        private readonly MessageBoxIcon icon;

        public MessageDialogForm(string text, string caption, MessageBoxButtons buttons,
            MessageBoxIcon icon, MessageBoxDefaultButton defaultButton) {
            this.icon = icon;
            Text = caption;
            EnableMinimizeButton = false;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9.75f);
            AutoScaleMode = AutoScaleMode.None;

            int textWidth = DialogWidth - TextLeft - Gutter;
            int maxTextHeight = Math.Max(120, Screen.PrimaryScreen.WorkingArea.Height / 2);
            const TextFormatFlags measure = TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix | TextFormatFlags.TextBoxControl;
            int measured = TextRenderer.MeasureText(text ?? string.Empty, Font, new Size(textWidth, 0), measure).Height;
            bool scrolling = measured > maxTextHeight;
            int bodyHeight = Math.Min(measured, maxTextHeight);

            var titleLabel = new Label {
                AutoSize = false,
                AutoEllipsis = true,
                UseMnemonic = false,
                Text = caption,
                Font = new Font("Segoe UI Semibold", 13f),
                Location = new Point(TextLeft, Gutter - 4),
                Size = new Size(textWidth, 28)
            };
            Controls.Add(titleLabel);

            int bodyTop = titleLabel.Bottom + 6;
            if (scrolling) {
                // Sehr lange Meldungen scrollbar und kopierbar anzeigen.
                Controls.Add(new ThemedTextBox {
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    BorderStyle = BorderStyle.None,
                    Text = text,
                    Location = new Point(TextLeft, bodyTop),
                    Size = new Size(textWidth, bodyHeight)
                });
            } else {
                Controls.Add(new Label {
                    AutoSize = false,
                    UseMnemonic = false,
                    Text = text,
                    Tag = "subtle",
                    Location = new Point(TextLeft, bodyTop),
                    Size = new Size(textWidth, bodyHeight + 4)
                });
            }

            int textBottom = Math.Max(bodyTop + bodyHeight + 4, Gutter + IconBoxSize);
            var actions = new FlowLayoutPanel {
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Location = new Point(Gutter, textBottom + 22),
                Size = new Size(DialogWidth - 2 * Gutter, ButtonHeight + 2),
                Margin = Padding.Empty
            };
            Controls.Add(actions);
            ClientSize = new Size(DialogWidth, actions.Bottom + Gutter - 4);

            AddButtons(actions, buttons, defaultButton);
        }

        private void AddButtons(FlowLayoutPanel actions, MessageBoxButtons buttons, MessageBoxDefaultButton defaultButton) {
            var definitions = new List<KeyValuePair<string, DialogResult>>();
            DialogResult escapeResult;
            switch (buttons) {
                case MessageBoxButtons.OKCancel:
                    definitions.Add(new KeyValuePair<string, DialogResult>("OK", DialogResult.OK));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Abbrechen", DialogResult.Cancel));
                    escapeResult = DialogResult.Cancel;
                    break;
                case MessageBoxButtons.YesNo:
                    definitions.Add(new KeyValuePair<string, DialogResult>("Ja", DialogResult.Yes));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Nein", DialogResult.No));
                    escapeResult = DialogResult.No;
                    break;
                case MessageBoxButtons.YesNoCancel:
                    definitions.Add(new KeyValuePair<string, DialogResult>("Ja", DialogResult.Yes));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Nein", DialogResult.No));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Abbrechen", DialogResult.Cancel));
                    escapeResult = DialogResult.Cancel;
                    break;
                case MessageBoxButtons.RetryCancel:
                    definitions.Add(new KeyValuePair<string, DialogResult>("Wiederholen", DialogResult.Retry));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Abbrechen", DialogResult.Cancel));
                    escapeResult = DialogResult.Cancel;
                    break;
                case MessageBoxButtons.AbortRetryIgnore:
                    definitions.Add(new KeyValuePair<string, DialogResult>("Abbrechen", DialogResult.Abort));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Wiederholen", DialogResult.Retry));
                    definitions.Add(new KeyValuePair<string, DialogResult>("Ignorieren", DialogResult.Ignore));
                    escapeResult = DialogResult.Abort;
                    break;
                default:
                    definitions.Add(new KeyValuePair<string, DialogResult>("OK", DialogResult.OK));
                    escapeResult = DialogResult.OK;
                    break;
            }

            int defaultIndex = 0;
            if (defaultButton == MessageBoxDefaultButton.Button2) defaultIndex = 1;
            if (defaultButton == MessageBoxDefaultButton.Button3) defaultIndex = 2;
            defaultIndex = Math.Min(defaultIndex, definitions.Count - 1);

            // FlowDirection.RightToLeft: letzter Eintrag steht ganz rechts, also umgekehrt hinzufügen.
            ThemedButton focusTarget = null;
            for (int i = definitions.Count - 1; i >= 0; i--) {
                var button = new ThemedButton {
                    Text = definitions[i].Key,
                    DialogResult = definitions[i].Value,
                    Kind = i == defaultIndex ? ButtonKind.Primary : ButtonKind.Secondary,
                    Size = new Size(110, ButtonHeight),
                    Margin = new Padding(10, 0, 0, 0),
                    TabIndex = i
                };
                actions.Controls.Add(button);
                if (i == defaultIndex) {
                    focusTarget = button;
                    AcceptButton = button;
                }
                if (definitions[i].Value == escapeResult) {
                    CancelButton = button;
                }
            }

            // Schließen über das X entspricht derselben Antwort wie die Esc-Taste.
            FormClosing += (s, e) => {
                if (DialogResult == DialogResult.None || DialogResult == DialogResult.Cancel) {
                    DialogResult = escapeResult;
                }
            };
            Shown += (s, e) => focusTarget?.Focus();
        }

        protected override void OnShown(EventArgs e) {
            base.OnShown(e);
            switch (icon) {
                case MessageBoxIcon.Error: SystemSounds.Hand.Play(); break;
                case MessageBoxIcon.Warning: SystemSounds.Exclamation.Play(); break;
                case MessageBoxIcon.Question: SystemSounds.Question.Play(); break;
                case MessageBoxIcon.Information: SystemSounds.Asterisk.Play(); break;
            }
        }

        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            if (ContentHost != null) {
                ContentHost.Paint += ContentHost_Paint;
            }
        }

        private void ContentHost_Paint(object sender, PaintEventArgs e) {
            IconKind kind;
            Tone tone;
            switch (icon) {
                case MessageBoxIcon.Error: kind = IconKind.ErrorCircle; tone = Tone.Error; break;
                case MessageBoxIcon.Warning: kind = IconKind.Warning; tone = Tone.Warning; break;
                case MessageBoxIcon.Question: kind = IconKind.Question; tone = Tone.Accent; break;
                case MessageBoxIcon.Information: kind = IconKind.Info; tone = Tone.Accent; break;
                default: return;
            }
            Icons.Draw(e.Graphics, kind, new RectangleF(Gutter, Gutter, IconBoxSize - 8, IconBoxSize - 8), AppTheme.ToneColor(tone), 1.6f);
        }
    }
}
