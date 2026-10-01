using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace WordLatexAddin
{
    /// <summary>User settings persisted below LocalAppData without modifying documents.</summary>
    [Serializable]
    public sealed class Settings
    {
        public bool AutomaticConversion { get; set; }
        public bool ShowCompletionMessages { get; set; }
        public int AutoDelayMilliseconds { get; set; }
        public string EquationSequenceName { get; set; }

        public Settings()
        {
            AutomaticConversion = false;
            ShowCompletionMessages = true;
            AutoDelayMilliseconds = 350;
            EquationSequenceName = "Equation";
        }

        public static string SettingsPath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordLatexVSTO", "settings.xml");
            }
        }

        public static Settings Load()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return new Settings();
                using (FileStream stream = File.OpenRead(SettingsPath))
                {
                    return (Settings)new XmlSerializer(typeof(Settings)).Deserialize(stream);
                }
            }
            catch
            {
                return new Settings();
            }
        }

        public void Save()
        {
            string directory = Path.GetDirectoryName(SettingsPath);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            string temporary = SettingsPath + ".tmp";
            using (FileStream stream = File.Create(temporary))
            {
                new XmlSerializer(typeof(Settings)).Serialize(stream, this);
            }
            if (File.Exists(SettingsPath)) File.Delete(SettingsPath);
            File.Move(temporary, SettingsPath);
        }

        public bool ShowDialog()
        {
            using (SettingsDialog dialog = new SettingsDialog(this))
            {
                return dialog.ShowDialog() == DialogResult.OK;
            }
        }
    }

    /// <summary>Small native settings dialog used by the Ribbon settings button.</summary>
    internal sealed class SettingsDialog : Form
    {
        private readonly Settings _settings;
        private readonly CheckBox _automatic;
        private readonly CheckBox _messages;
        private readonly NumericUpDown _delay;
        private readonly TextBox _sequence;

        public SettingsDialog(Settings settings)
        {
            _settings = settings;
            Text = "WordLatexVSTO 插件设置";
            Font = new Font("Microsoft YaHei UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(420, 230);

            _automatic = new CheckBox { Left = 22, Top = 22, Width = 360, Text = "启用输入结束后的自动转换", Checked = settings.AutomaticConversion };
            _messages = new CheckBox { Left = 22, Top = 55, Width = 360, Text = "批量转换后显示结果摘要", Checked = settings.ShowCompletionMessages };
            Label delayLabel = new Label { Left = 22, Top = 94, Width = 135, Text = "自动转换延迟（毫秒）" };
            _delay = new NumericUpDown { Left = 170, Top = 90, Width = 90, Minimum = 150, Maximum = 2000, Increment = 50, Value = Math.Max(150, Math.Min(2000, settings.AutoDelayMilliseconds)) };
            Label sequenceLabel = new Label { Left = 22, Top = 132, Width = 135, Text = "公式编号序列名" };
            _sequence = new TextBox { Left = 170, Top = 128, Width = 180, Text = settings.EquationSequenceName ?? "Equation" };

            Button ok = new Button { Left = 238, Top = 180, Width = 75, Text = "确定", DialogResult = DialogResult.OK };
            Button cancel = new Button { Left = 325, Top = 180, Width = 75, Text = "取消", DialogResult = DialogResult.Cancel };
            ok.Click += SaveValues;
            Controls.AddRange(new Control[] { _automatic, _messages, delayLabel, _delay, sequenceLabel, _sequence, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void SaveValues(object sender, EventArgs e)
        {
            _settings.AutomaticConversion = _automatic.Checked;
            _settings.ShowCompletionMessages = _messages.Checked;
            _settings.AutoDelayMilliseconds = Decimal.ToInt32(_delay.Value);
            _settings.EquationSequenceName = string.IsNullOrWhiteSpace(_sequence.Text) ? "Equation" : _sequence.Text.Trim();
            _settings.Save();
        }
    }
}
