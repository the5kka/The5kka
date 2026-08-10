using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace OjtExamPatch
{
    internal sealed class Question
    {
        public string SheetName = "";
        public string No = "";
        public string Type = "";
        public string Text = "";
        public string Answer = "";
        public double Score;
        public readonly List<Image> Images = new List<Image>();
    }

    internal sealed class Bank
    {
        public string Name = "";
        public string Department = "";
        public string Evaluator = "";
        public string JobName = "";
        public string Revision = "";
        public string IssueDate = "";
        public string ProductType = "";
        public readonly List<Question> Questions = new List<Question>();
        public int Vda { get { return Questions.Count(q => q.Type == "VDA"); } }
        public int Choice { get { return Questions.Count(q => q.Type == "\uAC1D\uAD00\uC2DD"); } }
        public int Subjective { get { return Questions.Count(q => q.Type == "\uC8FC\uAD00\uC2DD"); } }
        public int LegacyCommon;
    }

    internal sealed class ExamDoc
    {
        public Bank Bank;
        public readonly List<Question> Questions = new List<Question>();
        public bool ShowAnswers;
        public string UserName = "";
        public string EvalDate = "";
    }

    internal sealed class ExamModeSettings
    {
        public int VdaCount;
        public decimal VdaScore;
        public int ChoiceCount;
        public decimal ChoiceScore;
        public int SubjectiveCount;
        public decimal SubjectiveScore;
        public decimal TargetScore;

        public decimal Total
        {
            get
            {
                return VdaCount * VdaScore
                    + ChoiceCount * ChoiceScore
                    + SubjectiveCount * SubjectiveScore;
            }
        }

        public ExamModeSettings Clone()
        {
            return new ExamModeSettings
            {
                VdaCount = VdaCount,
                VdaScore = VdaScore,
                ChoiceCount = ChoiceCount,
                ChoiceScore = ChoiceScore,
                SubjectiveCount = SubjectiveCount,
                SubjectiveScore = SubjectiveScore,
                TargetScore = TargetScore
            };
        }
    }

    internal sealed class ExamSettings
    {
        public ExamModeSettings General;
        public ExamModeSettings Electric;

        public ExamSettings Clone()
        {
            return new ExamSettings
            {
                General = General.Clone(),
                Electric = Electric.Clone()
            };
        }
    }

    internal static class ExamSettingsStore
    {
        public static ExamSettings CreateDefault()
        {
            return new ExamSettings
            {
                General = new ExamModeSettings
                {
                    VdaCount = 0,
                    VdaScore = 2.5m,
                    ChoiceCount = 20,
                    ChoiceScore = 4m,
                    SubjectiveCount = 4,
                    SubjectiveScore = 5m,
                    TargetScore = 100m
                },
                Electric = new ExamModeSettings
                {
                    VdaCount = 2,
                    VdaScore = 2.5m,
                    ChoiceCount = 20,
                    ChoiceScore = 4m,
                    SubjectiveCount = 3,
                    SubjectiveScore = 5m,
                    TargetScore = 100m
                }
            };
        }

        public static ExamSettings Load()
        {
            ExamSettings settings = CreateDefault();
            try
            {
                string path = SettingsPath();
                if (!File.Exists(path)) return settings;
                XDocument doc = XDocument.Load(path);
                if (doc.Root == null) return settings;
                settings.General = ReadMode(doc.Root.Element("general"), settings.General, false);
                settings.Electric = ReadMode(doc.Root.Element("electric"), settings.Electric, true);
            }
            catch { }
            return settings;
        }

        public static void Save(ExamSettings settings)
        {
            if (settings == null || settings.General == null || settings.Electric == null)
                throw new InvalidOperationException("Invalid exam settings.");

            string path = SettingsPath();
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            string backup = path + ".bak";
            XDocument doc = new XDocument(
                new XElement("examSettings",
                    new XAttribute("version", "2"),
                    ModeElement("general", settings.General, false),
                    ModeElement("electric", settings.Electric, true)));
            doc.Save(temp);
            if (File.Exists(path))
            {
                if (File.Exists(backup)) File.Delete(backup);
                File.Replace(temp, path, backup, true);
                if (File.Exists(backup)) File.Delete(backup);
            }
            else
            {
                File.Move(temp, path);
            }
        }

        static string SettingsPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OJT_EXAM_MAKER",
                "exam_settings_v2.xml");
        }

        static XElement ModeElement(string name, ExamModeSettings mode, bool includeVda)
        {
            XElement element = new XElement(name,
                new XAttribute("choiceCount", mode.ChoiceCount),
                new XAttribute("choiceScore", Invariant(mode.ChoiceScore)),
                new XAttribute("subjectiveCount", mode.SubjectiveCount),
                new XAttribute("subjectiveScore", Invariant(mode.SubjectiveScore)),
                new XAttribute("targetScore", Invariant(mode.TargetScore)));
            if (includeVda)
            {
                element.Add(new XAttribute("vdaCount", mode.VdaCount));
                element.Add(new XAttribute("vdaScore", Invariant(mode.VdaScore)));
            }
            return element;
        }

        static ExamModeSettings ReadMode(XElement element, ExamModeSettings fallback, bool includeVda)
        {
            ExamModeSettings mode = fallback.Clone();
            if (element == null) return mode;
            mode.VdaCount = includeVda ? ReadInt(element, "vdaCount", mode.VdaCount, 0, 500) : 0;
            mode.VdaScore = includeVda ? ReadDecimal(element, "vdaScore", mode.VdaScore, 0m, 100m) : mode.VdaScore;
            mode.ChoiceCount = ReadInt(element, "choiceCount", mode.ChoiceCount, 0, 500);
            mode.ChoiceScore = ReadDecimal(element, "choiceScore", mode.ChoiceScore, 0m, 100m);
            mode.SubjectiveCount = ReadInt(element, "subjectiveCount", mode.SubjectiveCount, 0, 500);
            mode.SubjectiveScore = ReadDecimal(element, "subjectiveScore", mode.SubjectiveScore, 0m, 100m);
            mode.TargetScore = ReadDecimal(element, "targetScore", mode.TargetScore, 0m, 1000m);
            return mode;
        }

        static int ReadInt(XElement element, string name, int fallback, int min, int max)
        {
            int value;
            XAttribute attribute = element.Attribute(name);
            if (attribute == null || !int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return fallback;
            return Math.Max(min, Math.Min(max, value));
        }

        static decimal ReadDecimal(XElement element, string name, decimal fallback, decimal min, decimal max)
        {
            decimal value;
            XAttribute attribute = element.Attribute(name);
            if (attribute == null || !decimal.TryParse(attribute.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
                return fallback;
            return Math.Max(min, Math.Min(max, value));
        }

        static string Invariant(decimal value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }

    internal sealed class ExamModeEditor : UserControl
    {
        readonly bool includeVda;
        NumericUpDown vdaCount, vdaScore, choiceCount, choiceScore, subjectiveCount, subjectiveScore, targetScore;
        Label totalLabel;

        public ExamModeEditor(bool includeVda, ExamModeSettings settings)
        {
            this.includeVda = includeVda;
            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(18, 28, 46);
            BuildUi();
            LoadSettings(settings);
        }

        public ExamModeSettings ReadSettings()
        {
            return new ExamModeSettings
            {
                VdaCount = includeVda ? (int)vdaCount.Value : 0,
                VdaScore = includeVda ? vdaScore.Value : 2.5m,
                ChoiceCount = (int)choiceCount.Value,
                ChoiceScore = choiceScore.Value,
                SubjectiveCount = (int)subjectiveCount.Value,
                SubjectiveScore = subjectiveScore.Value,
                TargetScore = targetScore.Value
            };
        }

        void BuildUi()
        {
            int typeRows = includeVda ? 3 : 2;
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(18),
                ColumnCount = 3,
                RowCount = typeRows + 3,
                BackColor = Color.FromArgb(18, 28, 46)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            table.Controls.Add(Header(""), 0, 0);
            table.Controls.Add(Header("\uBB38\uC81C\uC218"), 1, 0);
            table.Controls.Add(Header("\uC810\uC218"), 2, 0);

            int row = 1;
            if (includeVda)
                AddTypeRow(table, row++, "VDA", out vdaCount, out vdaScore);
            AddTypeRow(table, row++, "\uAC1D\uAD00\uC2DD", out choiceCount, out choiceScore);
            AddTypeRow(table, row++, "\uC8FC\uAD00\uC2DD", out subjectiveCount, out subjectiveScore);

            table.Controls.Add(TypeLabel("\uBAA9\uD45C \uC810\uC218"), 0, row);
            targetScore = Number(0, 1000, 100, 0);
            table.Controls.Add(targetScore, 2, row++);
            totalLabel = new Label
            {
                Text = "TOTAL 100",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(32, 208, 132),
                BackColor = Color.FromArgb(16, 36, 58),
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                Margin = new Padding(0, 8, 0, 0)
            };
            table.Controls.Add(totalLabel, 0, row);
            table.SetColumnSpan(totalLabel, 3);
            Controls.Add(table);
        }

        void AddTypeRow(TableLayoutPanel table, int row, string name, out NumericUpDown count, out NumericUpDown score)
        {
            table.Controls.Add(TypeLabel(name), 0, row);
            count = Number(0, 500, 0, 0);
            score = Number(0, 100, 0, 1);
            table.Controls.Add(count, 1, row);
            table.Controls.Add(score, 2, row);
        }

        Label Header(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(176, 190, 215),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold)
            };
        }

        Label TypeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11, FontStyle.Bold)
            };
        }

        NumericUpDown Number(decimal min, decimal max, decimal value, int decimals)
        {
            var number = new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = value,
                DecimalPlaces = decimals,
                Increment = decimals == 0 ? 1m : .5m,
                TextAlign = HorizontalAlignment.Center,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(8, 17, 31),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Margin = new Padding(6)
            };
            number.ValueChanged += delegate { UpdateTotal(); };
            return number;
        }

        void LoadSettings(ExamModeSettings settings)
        {
            if (settings == null) return;
            if (includeVda)
            {
                SetValue(vdaCount, settings.VdaCount);
                SetValue(vdaScore, settings.VdaScore);
            }
            SetValue(choiceCount, settings.ChoiceCount);
            SetValue(choiceScore, settings.ChoiceScore);
            SetValue(subjectiveCount, settings.SubjectiveCount);
            SetValue(subjectiveScore, settings.SubjectiveScore);
            SetValue(targetScore, settings.TargetScore);
            UpdateTotal();
        }

        static void SetValue(NumericUpDown number, decimal value)
        {
            number.Value = Math.Max(number.Minimum, Math.Min(number.Maximum, value));
        }

        void UpdateTotal()
        {
            if (totalLabel == null || choiceCount == null || subjectiveCount == null || targetScore == null)
                return;
            ExamModeSettings settings = ReadSettings();
            totalLabel.Text = "TOTAL " + settings.Total.ToString("0.##");
            totalLabel.ForeColor = settings.Total == settings.TargetScore
                ? Color.FromArgb(32, 208, 132)
                : Color.FromArgb(246, 180, 75);
        }
    }

    internal static class SettingsAccess
    {
        const string Salt = "OJT_EXAM_MAKER_V2|";
        const string PasswordHash = "E735E3D89CD4BB607A7115441D59D28F37384B9327B34302F298A622F9F3B233";

        public static bool Verify(string password)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(Salt + (password ?? "")));
                string hash = BitConverter.ToString(bytes).Replace("-", "");
                return string.Equals(hash, PasswordHash, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    internal sealed class SettingsPasswordForm : Form
    {
        readonly TextBox passwordBox;

        public SettingsPasswordForm()
        {
            Text = "\uC124\uC815 \uBE44\uBC00\uBC88\uD638";
            Width = 430;
            Height = 260;
            MinimumSize = new Size(430, 260);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(11, 18, 32);
            Font = new Font("\uB9D1\uC740 \uACE0\uB515", 10);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 20, 24, 18),
                ColumnCount = 1,
                RowCount = 3
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            root.Controls.Add(new Label
            {
                Text = "\uC124\uC815\uC744 \uBCC0\uACBD\uD558\uB824\uBA74\n\uAD00\uB9AC\uC790 \uBE44\uBC00\uBC88\uD638\uB97C \uC785\uB825\uD558\uC138\uC694.",
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11, FontStyle.Bold)
            }, 0, 0);

            passwordBox = new TextBox
            {
                Dock = DockStyle.Fill,
                UseSystemPasswordChar = true,
                MaxLength = 20,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.FromArgb(8, 17, 31),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                Margin = new Padding(26, 4, 26, 4)
            };
            root.Controls.Add(passwordBox, 0, 1);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 12, 0, 0)
            };
            var confirm = PasswordButton("\uD655\uC778");
            var cancel = PasswordButton("\uCDE8\uC18C");
            confirm.Click += Confirm;
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(confirm);
            buttons.Controls.Add(cancel);
            root.Controls.Add(buttons, 0, 2);
            AcceptButton = confirm;
            CancelButton = cancel;
            Shown += delegate { passwordBox.Focus(); };
        }

        Button PasswordButton(string text)
        {
            return new Button
            {
                Text = text,
                Width = 100,
                Height = 32,
                BackColor = Color.FromArgb(30, 45, 72),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold),
                Margin = new Padding(8, 0, 0, 0)
            };
        }

        void Confirm(object sender, EventArgs e)
        {
            if (SettingsAccess.Verify(passwordBox.Text))
            {
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            MessageBox.Show(
                this,
                "\uBE44\uBC00\uBC88\uD638\uAC00 \uB9DE\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4.",
                "\uC124\uC815 \uC7A0\uAE08",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            passwordBox.Clear();
            passwordBox.Focus();
        }
    }

    internal sealed class ExamSettingsForm : Form
    {
        readonly ExamModeEditor generalEditor;
        readonly ExamModeEditor electricEditor;
        public ExamSettings Value { get; private set; }

        public ExamSettingsForm(ExamSettings settings)
        {
            ExamSettings source = (settings ?? ExamSettingsStore.CreateDefault()).Clone();
            Text = "\uC2DC\uD5D8 \uAE30\uBCF8 \uC124\uC815";
            Width = 650;
            Height = 460;
            MinimumSize = new Size(650, 460);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(11, 18, 32);
            Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(14) };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            Controls.Add(root);

            var tabs = new TabControl { Dock = DockStyle.Fill };
            var generalTab = new TabPage("\uC2E0\uC785\u00B7\uC77C\uBC18") { BackColor = Color.FromArgb(18, 28, 46) };
            var electricTab = new TabPage("\uC804\uC7A5") { BackColor = Color.FromArgb(18, 28, 46) };
            generalEditor = new ExamModeEditor(false, source.General);
            electricEditor = new ExamModeEditor(true, source.Electric);
            generalTab.Controls.Add(generalEditor);
            electricTab.Controls.Add(electricEditor);
            tabs.TabPages.Add(generalTab);
            tabs.TabPages.Add(electricTab);
            root.Controls.Add(tabs, 0, 0);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 8, 0, 0)
            };
            var save = DialogButton("\uC800\uC7A5");
            var cancel = DialogButton("\uCDE8\uC18C");
            save.Click += SaveAndClose;
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);
            root.Controls.Add(buttons, 0, 1);
            AcceptButton = save;
            CancelButton = cancel;
        }

        Button DialogButton(string text)
        {
            return new Button
            {
                Text = text,
                Width = 100,
                Height = 32,
                BackColor = Color.FromArgb(30, 45, 72),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold),
                Margin = new Padding(8, 0, 0, 0)
            };
        }

        void SaveAndClose(object sender, EventArgs e)
        {
            ExamModeSettings general = generalEditor.ReadSettings();
            ExamModeSettings electric = electricEditor.ReadSettings();
            if (general.Total != general.TargetScore || electric.Total != electric.TargetScore)
            {
                MessageBox.Show(
                    this,
                    "\uC124\uC815 \uD569\uACC4\uC640 \uBAA9\uD45C \uC810\uC218\uAC00 \uB2E4\uB985\uB2C8\uB2E4.\n\uB450 \uD0ED\uC758 TOTAL\uC744 \uD655\uC778\uD558\uC138\uC694.",
                    "\uC124\uC815 \uD655\uC778",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            Value = new ExamSettings { General = general, Electric = electric };
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    internal sealed class SheetPart
    {
        public string Name = "";
        public string Path = "";
        public string State = "";
    }

    internal sealed class Header
    {
        public int Row, NoCol, TypeCol, QuestionCol, AnswerCol, ScoreCol;
        public readonly List<int> ChoiceCols = new List<int>();
    }

    internal sealed class Reader
    {
        static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly XNamespace PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        static readonly XNamespace DrawNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public List<Bank> Load(string file)
        {
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = SharedStrings(zip);
                var sheets = Sheets(zip);
                var settings = Settings(zip, sheets, shared);
                var banks = new List<Bank>();
                var used = new Dictionary<string, int>();

                foreach (var sheet in sheets)
                {
                    string sheetName = sheet.Name.Trim();
                    if (sheet.State == "hidden" || sheet.State == "veryHidden")
                        continue;
                    if (sheetName == "\uC2DC\uD5D8 SETTING" || sheetName == "\uC2DC\uD5D8\uC9C0" || sheetName == "\uB2F5\uC548\uC9C0")
                        continue;
                    var cells = Cells(zip, sheet.Path, shared);
                    var header = FindHeader(cells);
                    if (header == null)
                        continue;
                    var images = ImagesByRow(zip, sheet.Path, header.QuestionCol, header.AnswerCol);
                    var bank = new Bank { Name = sheetName };
                    ApplyMeta(bank, settings);
                    int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
                    for (int row = header.Row + 1; row <= maxRow; row++)
                    {
                        string no = Clean(Get(cells, row, header.NoCol));
                        string questionText = BuildQuestion(cells, row, header);
                        if (no.Length == 0 || questionText.Length == 0)
                            continue;
                        string type = Clean(Get(cells, row, header.TypeCol));
                        if (type == "\uACF5\uD1B5")
                        {
                            bank.LegacyCommon++;
                            continue;
                        }
                        if (type != "VDA" && type != "\uAC1D\uAD00\uC2DD" && type != "\uC8FC\uAD00\uC2DD")
                            continue;
                        double score;
                        if (!double.TryParse(Clean(Get(cells, row, header.ScoreCol)), out score))
                            score = type == "VDA" ? 2.5 : type == "\uAC1D\uAD00\uC2DD" ? 4.0 : 5.0;
                        var q = new Question
                        {
                            SheetName = sheetName,
                            No = no,
                            Type = type,
                            Text = questionText,
                            Answer = Clean(Get(cells, row, header.AnswerCol)),
                            Score = score
                        };
                        List<Image> rowImages;
                        if (images.TryGetValue(row, out rowImages))
                            q.Images.AddRange(rowImages);
                        bank.Questions.Add(q);
                    }
                    if (bank.Questions.Count > 0)
                    {
                        used[sheetName] = used.ContainsKey(sheetName) ? used[sheetName] + 1 : 1;
                        if (used[sheetName] > 1)
                            bank.Name = sheetName + " (" + used[sheetName] + ")";
                        banks.Add(bank);
                    }
                }
                return banks.OrderBy(b => b.Name).ToList();
            }
        }

        static Dictionary<string, string> Settings(ZipArchive zip, List<SheetPart> sheets, List<string> shared)
        {
            var setting = sheets.FirstOrDefault(s => s.Name.Trim() == "\uC2DC\uD5D8 SETTING");
            var result = new Dictionary<string, string>();
            if (setting == null)
                return result;
            var cells = Cells(zip, setting.Path, shared);
            int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
            for (int row = 3; row <= maxRow; row++)
            {
                string key = Clean(Get(cells, row, 11));
                if (key.Length == 0) continue;
                result[key.Trim()] = string.Join("\t", new[] {
                    Clean(Get(cells,row,12)), Clean(Get(cells,row,13)), Clean(Get(cells,row,14)),
                    Clean(Get(cells,row,15)), Clean(Get(cells,row,16)), Clean(Get(cells,row,17))
                });
            }
            return result;
        }

        static void ApplyMeta(Bank b, Dictionary<string, string> settings)
        {
            string packed;
            if (settings.TryGetValue(b.Name.Trim(), out packed))
            {
                var p = packed.Split('\t');
                if (p.Length > 0 && p[0].Length > 0) b.Department = p[0];
                if (p.Length > 1 && p[1].Length > 0) b.Evaluator = p[1];
                if (p.Length > 2 && p[2].Length > 0) b.JobName = p[2];
                if (p.Length > 3 && p[3].Length > 0) b.Revision = p[3];
                if (p.Length > 4 && p[4].Length > 0) b.IssueDate = p[4];
                if (p.Length > 5 && p[5].Length > 0) b.ProductType = p[5];
            }
        }

        static string BuildQuestion(Dictionary<Tuple<int, int>, string> cells, int row, Header h)
        {
            string text = Clean(Get(cells, row, h.QuestionCol));
            if (h.ChoiceCols.Count == 0)
                return text;
            string[] marks = { "\u2460", "\u2461", "\u2462", "\u2463", "\u2464", "\u2465" };
            var parts = new List<string>();
            for (int i = 0; i < h.ChoiceCols.Count && i < marks.Length; i++)
            {
                string value = Clean(Get(cells, row, h.ChoiceCols[i]));
                if (value.Length > 0)
                    parts.Add(marks[i] + " " + value);
            }
            return parts.Count == 0 ? text : text + "\n\n" + string.Join("\n", parts);
        }

        static Header FindHeader(Dictionary<Tuple<int, int>, string> cells)
        {
            int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
            int maxCol = cells.Keys.Select(k => k.Item2).DefaultIfEmpty(0).Max();
            for (int row = 1; row <= Math.Min(maxRow, 12); row++)
            {
                var h = new Dictionary<string, int>();
                for (int col = 1; col <= maxCol; col++)
                {
                    string key = Norm(Get(cells, row, col));
                    if (key.Length > 0 && !h.ContainsKey(key))
                        h[key] = col;
                }
                if (!h.ContainsKey("no") || !h.ContainsKey("\uBB38\uC81C") || !h.ContainsKey("\uBB38\uC81C\uC720\uD615"))
                    continue;
                string answerKey = h.ContainsKey("\uB2F5\uC548") ? "\uB2F5\uC548" : h.ContainsKey("\uC815\uB2F5") ? "\uC815\uB2F5" : null;
                if (answerKey == null)
                    continue;
                var header = new Header
                {
                    Row = row,
                    NoCol = h["no"],
                    TypeCol = h["\uBB38\uC81C\uC720\uD615"],
                    QuestionCol = h["\uBB38\uC81C"],
                    AnswerCol = h[answerKey],
                    ScoreCol = h.ContainsKey("\uC810\uC218") ? h["\uC810\uC218"] : 0
                };
                for (int n = 1; n <= 6; n++)
                    if (h.ContainsKey(n.ToString()))
                        header.ChoiceCols.Add(h[n.ToString()]);
                return header;
            }
            return null;
        }

        static List<string> SharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var si in doc.Descendants(MainNs + "si"))
                    list.Add(string.Concat(si.Descendants(MainNs + "t").Select(t => (string)t)));
            }
            return list;
        }

        static List<SheetPart> Sheets(ZipArchive zip)
        {
            var rels = new Dictionary<string, string>();
            using (var s = zip.GetEntry("xl/_rels/workbook.xml.rels").Open())
            {
                var doc = XDocument.Load(s);
                foreach (var r in doc.Root.Elements(PkgRelNs + "Relationship"))
                    rels[(string)r.Attribute("Id")] = ResolvePart("xl/workbook.xml", (string)r.Attribute("Target"));
            }
            var sheets = new List<SheetPart>();
            using (var s = zip.GetEntry("xl/workbook.xml").Open())
            {
                var doc = XDocument.Load(s);
                foreach (var sh in doc.Descendants(MainNs + "sheet"))
                {
                    string id = (string)sh.Attribute(RelNs + "id");
                    if (id != null && rels.ContainsKey(id))
                        sheets.Add(new SheetPart { Name = (string)sh.Attribute("name") ?? "", Path = rels[id], State = (string)sh.Attribute("state") ?? "" });
                }
            }
            return sheets;
        }

        static Dictionary<Tuple<int, int>, string> Cells(ZipArchive zip, string sheetPath, List<string> shared)
        {
            var cells = new Dictionary<Tuple<int, int>, string>();
            var entry = zip.GetEntry(sheetPath);
            if (entry == null) return cells;
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var c in doc.Descendants(MainNs + "c"))
                {
                    int row, col; ParseRef((string)c.Attribute("r"), out row, out col);
                    string type = (string)c.Attribute("t");
                    string value;
                    if (type == "inlineStr")
                        value = string.Concat(c.Descendants(MainNs + "t").Select(t => (string)t));
                    else
                        value = (string)c.Element(MainNs + "v") ?? "";
                    if (type == "s")
                    {
                        int idx;
                        if (int.TryParse(value, out idx) && idx >= 0 && idx < shared.Count)
                            value = shared[idx];
                    }
                    cells[Tuple.Create(row, col)] = value;
                }
            }
            return cells;
        }

        static Dictionary<int, List<Image>> ImagesByRow(ZipArchive zip, string sheetPath, int questionCol, int answerCol)
        {
            var staged = new Dictionary<int, List<Tuple<int, long, Image>>>();
            var sheetRels = zip.GetEntry(PathToRels(sheetPath));
            if (sheetRels == null) return new Dictionary<int, List<Image>>();
            string drawingPath = null;
            using (var s = sheetRels.Open())
            {
                var doc = XDocument.Load(s);
                var rel = doc.Root.Elements(PkgRelNs + "Relationship").FirstOrDefault(r => ((string)r.Attribute("Type") ?? "").EndsWith("/drawing"));
                if (rel != null)
                    drawingPath = ResolvePart(sheetPath, (string)rel.Attribute("Target"));
            }
            if (drawingPath == null || zip.GetEntry(drawingPath) == null) return new Dictionary<int, List<Image>>();
            var imgRels = new Dictionary<string, string>();
            var drawingRels = zip.GetEntry(PathToRels(drawingPath));
            if (drawingRels != null)
            {
                using (var s = drawingRels.Open())
                {
                    var doc = XDocument.Load(s);
                    foreach (var r in doc.Root.Elements(PkgRelNs + "Relationship"))
                        imgRels[(string)r.Attribute("Id")] = ResolvePart(drawingPath, (string)r.Attribute("Target"));
                }
            }
            using (var s = zip.GetEntry(drawingPath).Open())
            {
                var doc = XDocument.Load(s);
                foreach (var anchor in doc.Root.Elements())
                {
                    var from = anchor.Element(DrawNs + "from");
                    var blip = anchor.Descendants(A + "blip").FirstOrDefault();
                    if (from == null) continue;
                    int row = ToInt((string)from.Element(DrawNs + "row")) + 1;
                    int col = ToInt((string)from.Element(DrawNs + "col")) + 1;
                    long colOff = ToLong((string)from.Element(DrawNs + "colOff"));
                    if (col < questionCol || (answerCol > questionCol && col >= answerCol))
                        continue;
                    var group = anchor.Element(DrawNs + "grpSp");
                    if (blip == null && group != null)
                    {
                        Image shapeImage = RenderGroupShape(group);
                        if (shapeImage != null)
                            StageImage(staged, row, col, colOff, shapeImage);
                        continue;
                    }
                    if (blip == null) continue;
                    string rid = (string)blip.Attribute(RelNs + "embed");
                    string imgPath;
                    if (rid == null || !imgRels.TryGetValue(rid, out imgPath)) continue;
                    var imgEntry = zip.GetEntry(imgPath);
                    if (imgEntry == null) continue;
                    using (var ims = imgEntry.Open())
                    using (var ms = new MemoryStream())
                    {
                        ims.CopyTo(ms);
                        ms.Position = 0;
                        Image img = Image.FromStream(ms);
                        StageImage(staged, row, col, colOff, new Bitmap(img));
                    }
                }
            }
            return staged.ToDictionary(
                pair => pair.Key,
                pair => pair.Value.OrderBy(item => item.Item1).ThenBy(item => item.Item2).Select(item => item.Item3).ToList());
        }
        static void StageImage(Dictionary<int, List<Tuple<int, long, Image>>> staged, int row, int col, long colOff, Image image)
        {
            if (!staged.ContainsKey(row)) staged[row] = new List<Tuple<int, long, Image>>();
            staged[row].Add(Tuple.Create(col, colOff, image));
        }
        static Image RenderGroupShape(XElement group)
        {
            var grpPr = group.Element(DrawNs + "grpSpPr");
            var xfrm = grpPr == null ? null : grpPr.Element(A + "xfrm");
            if (xfrm == null) return null;
            var chOff = xfrm.Element(A + "chOff");
            var chExt = xfrm.Element(A + "chExt");
            double ox = ToDouble(chOff == null ? null : (string)chOff.Attribute("x"));
            double oy = ToDouble(chOff == null ? null : (string)chOff.Attribute("y"));
            double ew = Math.Max(1, ToDouble(chExt == null ? null : (string)chExt.Attribute("cx")));
            double eh = Math.Max(1, ToDouble(chExt == null ? null : (string)chExt.Attribute("cy")));
            int width = 96, height = 76;
            var bmp = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                foreach (var sp in group.Elements(DrawNs + "sp"))
                {
                    var spPr = sp.Element(DrawNs + "spPr");
                    var sx = spPr == null ? null : spPr.Element(A + "xfrm");
                    if (sx == null) continue;
                    var off = sx.Element(A + "off");
                    var ext = sx.Element(A + "ext");
                    float x = (float)((ToDouble(off == null ? null : (string)off.Attribute("x")) - ox) / ew * width);
                    float y = (float)((ToDouble(off == null ? null : (string)off.Attribute("y")) - oy) / eh * height);
                    float w = Math.Max(1, (float)(ToDouble(ext == null ? null : (string)ext.Attribute("cx")) / ew * width));
                    float h = Math.Max(1, (float)(ToDouble(ext == null ? null : (string)ext.Attribute("cy")) / eh * height));
                    using (var b = new SolidBrush(ShapeColor(sp, Color.FromArgb(21, 79, 99))))
                        g.FillRectangle(b, x, y, w, h);
                }
                foreach (var sp in group.Elements(DrawNs + "cxnSp"))
                {
                    var spPr = sp.Element(DrawNs + "spPr");
                    var sx = spPr == null ? null : spPr.Element(A + "xfrm");
                    if (sx == null) continue;
                    var off = sx.Element(A + "off");
                    var ext = sx.Element(A + "ext");
                    float x = (float)((ToDouble(off == null ? null : (string)off.Attribute("x")) - ox) / ew * width);
                    float y = (float)((ToDouble(off == null ? null : (string)off.Attribute("y")) - oy) / eh * height);
                    float w = (float)(ToDouble(ext == null ? null : (string)ext.Attribute("cx")) / ew * width);
                    float h = (float)(ToDouble(ext == null ? null : (string)ext.Attribute("cy")) / eh * height);
                    using (var p = new Pen(ShapeColor(sp, Color.Red), 1.6f))
                        g.DrawLine(p, x, y, x + w, y + h);
                }
            }
            return bmp;
        }
        static Color ShapeColor(XElement el, Color fallback)
        {
            var srgb = el.Descendants(A + "srgbClr").FirstOrDefault();
            string val = srgb == null ? null : (string)srgb.Attribute("val");
            int rgb;
            if (!string.IsNullOrEmpty(val) && int.TryParse(val, System.Globalization.NumberStyles.HexNumber, null, out rgb))
                return Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
            return fallback;
        }

        static string Clean(string s)
        {
            return (s ?? "").Replace("\r\n", "\n").Replace("\r", "\n").Replace("\u00a0", " ").Trim();
        }
        static string Norm(string s) { return Regex.Replace(Clean(s), @"\s+", "").ToLowerInvariant(); }
        static string Get(Dictionary<Tuple<int, int>, string> cells, int row, int col)
        {
            if (row <= 0 || col <= 0) return "";
            string value; return cells.TryGetValue(Tuple.Create(row, col), out value) ? value : "";
        }
        static int ToInt(string s) { int n; return int.TryParse(s, out n) ? n : 0; }
        static long ToLong(string s) { long n; return long.TryParse(s, out n) ? n : 0; }
        static double ToDouble(string s) { double n; return double.TryParse(s, out n) ? n : 0; }
        static void ParseRef(string r, out int row, out int col)
        {
            row = 0; col = 0;
            foreach (char ch in r ?? "")
            {
                if (char.IsLetter(ch)) col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                else if (char.IsDigit(ch)) row = row * 10 + (ch - '0');
            }
        }
        static string ResolvePart(string basePart, string target)
        {
            target = (target ?? "").Replace('\\', '/');
            if (target.StartsWith("/")) return target.TrimStart('/');
            var parts = basePart.Split('/').ToList();
            parts.RemoveAt(parts.Count - 1);
            foreach (string p in target.Split('/'))
            {
                if (p == "." || p.Length == 0) continue;
                if (p == ".." && parts.Count > 0) parts.RemoveAt(parts.Count - 1);
                else if (p != "..") parts.Add(p);
            }
            return string.Join("/", parts);
        }
        static string PathToRels(string part)
        {
            int slash = part.LastIndexOf('/');
            string dir = slash >= 0 ? part.Substring(0, slash + 1) : "";
            string file = slash >= 0 ? part.Substring(slash + 1) : part;
            return dir + "_rels/" + file + ".rels";
        }
    }

    internal sealed class Preview : Control
    {
        const int PrintWrapUnits = 108;
        const int ScoreWrapUnits = 94;
        const int QuestionGapUnits = 14;
        public ExamDoc Doc;
        public int PageIndex;
        public float Zoom = 1f;
        Point pan;
        Point dragStart;
        bool dragging;
        readonly List<int> starts = new List<int>();
        public Preview()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Color.FromArgb(7, 16, 29);
        }
        public void ZoomBy(float factor)
        {
            Zoom = Math.Max(0.7f, Math.Min(2.4f, Zoom * factor));
            Invalidate();
        }
        public void Fit()
        {
            Zoom = 1f;
            pan = Point.Empty;
            Invalidate();
        }
        public int QuestionPageCount { get { BuildPages(); return Math.Max(1, starts.Count); } }
        public int PageCount { get { BuildPages(); return Doc == null ? 1 : Math.Max(1, starts.Count) + 1; } }
        public void MovePage(int delta) { BuildPages(); PageIndex = Math.Max(0, Math.Min(PageIndex + delta, PageCount - 1)); Invalidate(); }
        protected override void OnMouseWheel(MouseEventArgs e) { ZoomBy(e.Delta > 0 ? 1.12f : 1f / 1.12f); }
        protected override void OnMouseDown(MouseEventArgs e) { if (Zoom > 1f) { dragging = true; dragStart = e.Location; Cursor = Cursors.SizeAll; } base.OnMouseDown(e); }
        protected override void OnMouseMove(MouseEventArgs e) { if (dragging) { pan.X += e.X - dragStart.X; pan.Y += e.Y - dragStart.Y; dragStart = e.Location; Invalidate(); } base.OnMouseMove(e); }
        protected override void OnMouseUp(MouseEventArgs e) { dragging = false; Cursor = Cursors.Default; base.OnMouseUp(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.None;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.Half;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            if (Doc == null)
            {
                using (var b = new SolidBrush(Color.FromArgb(188, 209, 255)))
                using (var f = new Font("\uB9D1\uC740 \uACE0\uB515", 13, FontStyle.Bold))
                    e.Graphics.DrawString("\uBB38\uD56D \uC870\uAC74\uC744 \uC120\uD0DD\uD558\uACE0 [\uB79C\uB364 \uBBF8\uB9AC\uBCF4\uAE30]\uB97C \uB204\uB974\uC138\uC694.", f, b, 30, 32);
                return;
            }
            Rectangle page = PageRect(ClientRectangle);
            using (var printPage = new Bitmap(827, 1169))
            using (var printGraphics = Graphics.FromImage(printPage))
            {
                printGraphics.Clear(Color.White);
                printGraphics.SmoothingMode = SmoothingMode.None;
                printGraphics.PixelOffsetMode = PixelOffsetMode.Half;
                printGraphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle printContent = new Rectangle(18, 18, 791, 1133);
                if (PageIndex >= QuestionPageCount)
                    DrawAnswerPage(printGraphics, printContent);
                else
                    DrawPage(printGraphics, printContent, PageIndex, Doc.ShowAnswers);

                e.Graphics.CompositingQuality = CompositingQuality.HighQuality;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                e.Graphics.DrawImage(printPage, page);
            }
            e.Graphics.DrawRectangle(Pens.LightGray, page);
        }
        public void DrawPrint(Graphics g, Rectangle bounds, int page, bool answers) { DrawPage(g, bounds, page, answers); }
        public void DrawAnswerPrint(Graphics g, Rectangle bounds) { DrawAnswerPage(g, bounds); }
        Rectangle PageRect(Rectangle bounds)
        {
            float ratio = 793f / 1122f;
            int w = bounds.Width - 28;
            int h = (int)(w / ratio);
            int maxH = bounds.Height - 28;
            if (h > maxH) { h = maxH; w = (int)(h * ratio); }
            w = (int)(w * Zoom);
            h = (int)(h * Zoom);
            return new Rectangle(bounds.X + (bounds.Width - w) / 2 + pan.X, bounds.Y + (bounds.Height - h) / 2 + pan.Y, w, h);
        }
        Rectangle PrintContentRect(Rectangle page)
        {
            int mx = Math.Max(1, (int)Math.Round(page.Width * 18f / 827f));
            int my = Math.Max(1, (int)Math.Round(page.Height * 18f / 1169f));
            return new Rectangle(page.Left + mx, page.Top + my, page.Width - mx * 2, page.Height - my * 2);
        }
        void DrawPageFit(Graphics g, Rectangle bounds, int pageIndex, bool answers)
        {
            const float designW = 1122f;
            const float designH = 793f;
            float scale = Math.Min(bounds.Width / designW, bounds.Height / designH);
            int drawW = (int)Math.Round(designW * scale);
            int drawH = (int)Math.Round(designH * scale);
            int x = bounds.Left + (bounds.Width - drawW) / 2;
            int y = bounds.Top;
            GraphicsState state = g.Save();
            g.TranslateTransform(x, y);
            g.ScaleTransform(scale, scale);
            DrawPage(g, new Rectangle(0, 0, (int)designW, (int)designH), pageIndex, answers);
            g.Restore(state);
        }
        void BuildPages()
        {
            starts.Clear();
            if (Doc == null || Doc.Questions.Count == 0) return;
            starts.Add(0);
            int used = 0;
            int limit = 555;
            for (int i = 0; i < Doc.Questions.Count; i++)
            {
                int h = QHeight2(Doc.Questions[i]);
                int gap = used > 0 ? QuestionGapUnits : 0;
                if (i > 0 && used + gap + h > limit)
                {
                    starts.Add(i);
                    used = h;
                    limit = 760;
                }
                else
                {
                    used += gap + h;
                }
            }
            int maxPage = Math.Max(1, starts.Count) + 1;
            if (PageIndex >= maxPage) PageIndex = maxPage - 1;
        }
        int EndIndex(int page) { return page + 1 < starts.Count ? starts[page + 1] : Doc.Questions.Count; }
        public static string AnswerMark(Question q)
        {
            if (q.Type == "\uAC1D\uAD00\uC2DD" || q.Type == "VDA")
            {
                var m = Regex.Match(q.Answer ?? "", "[①②③④⑤⑥1-9]");
                if (m.Success)
                {
                    string circles = "①②③④⑤⑥";
                    int idx = circles.IndexOf(m.Value, StringComparison.Ordinal);
                    if (idx >= 0) return circles[idx].ToString();
                    int n;
                    if (int.TryParse(m.Value, out n) && n >= 1 && n <= circles.Length) return circles[n - 1].ToString();
                }
            }
            return q.Answer;
        }
        static int QHeight(Question q)
        {
            int lines = Math.Max(1, q.Text.Count(c => c == '\n') + 1);
            int textLines = Math.Max(lines, q.Text.Length / 92 + 1);
            int image = q.Images.Count == 0 ? 0 : 112;
            int subj = q.Type == "\uC8FC\uAD00\uC2DD" ? 28 : 0;
            return Math.Max(q.Images.Count > 0 ? 150 : 42, textLines * 14 + image + subj + 12);
        }
        static int QHeight2(Question q)
        {
            string display = "88.  " + FormatQuestionText(q);
            if (q.Images.Count > 0)
            {
                string mainText = CompactVisibleLines(StripImageChoiceLines(display));
                string choiceText = CompactVisibleLines(ImageChoiceLines(display));
                int mainLines = Math.Max(1, VisibleWrappedLineCount(mainText, ScoreWrapUnits));
                List<string> choices = ImageChoiceLineList(choiceText, q.Images.Count);
                bool choicesBelongToImages = ChoicesBelongToImages(q.Images.Count, choices.Count);
                int choiceLines = choicesBelongToImages
                    ? (choices.Take(Math.Min(q.Images.Count, choices.Count)).Any(line => ImageChoiceCaption(line).Length > 0) ? 1 : 0)
                        + Math.Max(0, choices.Count - Math.Min(q.Images.Count, choices.Count))
                    : choices.Count;
                bool wideMultiImage = HasWideMultiImages(q);
                int image = q.Images.Count == 1 ? 75 : wideMultiImage ? 72 : 92;
                int minimumHeight = wideMultiImage ? 105 : 125;
                return Math.Max(minimumHeight, mainLines * 13 + choiceLines * 13 + image + 12);
            }
            List<string> wrapped = WrapPrintLines(display, ScoreWrapUnits);
            int textLines = Math.Max(1, wrapped.Count(line => line.Trim().Length > 0));
            int blankLines = wrapped.Count - textLines;
            bool subjType = q.Type == "\uC8FC\uAD00\uC2DD";
            if (subjType)
                return Math.Max(70, textLines * 15 + blankLines * 12 + 8);
            return Math.Max(32, textLines * 12 + blankLines * 10 + 6);
        }
        void DrawPage(Graphics g, Rectangle page, int pageIndex, bool answers)
        {
            BuildPages();
            float sx = page.Width / 1122f, sy = page.Height / 793f;
            Func<float, int> X = v => page.Left + (int)(v * sx);
            Func<float, int> Y = v => page.Top + (int)(v * sy);
            Func<float, int> W = v => Math.Max(1, (int)(v * sx));
            Func<float, int> H = v => Math.Max(1, (int)(v * sy));
            g.SmoothingMode = SmoothingMode.None;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            using (var pen = new Pen(Color.Black, Math.Max(1, W(1))))
            using (var text = new SolidBrush(Color.Black))
            using (var title = new Font("\uB9D1\uC740 \uACE0\uB515", Math.Max(21, W(36)), FontStyle.Bold))
            using (var font = new Font("\uB9D1\uC740 \uACE0\uB515", Math.Max(9, W(10)), FontStyle.Regular))
            using (var bold = new Font("\uB9D1\uC740 \uACE0\uB515", Math.Max(9, W(10)), FontStyle.Bold))
            {
                int y;
                if (pageIndex == 0)
                {
                    DrawHeader(g, pen, title, font, bold, X, Y, W, H);
                    y = Y(214);
                }
                else
                {
                    y = Y(18);
                }
                y += H(4);
                int start = starts.Count == 0 ? 0 : starts[Math.Min(pageIndex, starts.Count - 1)];
                int end = EndIndex(Math.Min(pageIndex, PageCount - 1));
                for (int i = start; i < end; i++)
                {
                    if (i > start) y += H(QuestionGapUnits);
                    Question q = Doc.Questions[i];
                    int rh = H(QHeight2(q));
                    QuestionCell2(g, q, i + 1, font, bold, text, new Rectangle(X(48), y, W(1058), rh), answers);
                    y += rh;
                }
                DrawFooter(g, font, page);
            }
        }
        static void DrawFooter(Graphics g, Font font, Rectangle page)
        {
            using (var footerFont = new Font("\uB9D1\uC740 \uACE0\uB515", Math.Max(7, page.Width / 135), FontStyle.Regular))
            using (var b = new SolidBrush(Color.Black))
            {
                var left = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
                var center = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                var right = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
                int h = Math.Max(14, footerFont.Height + 4);
                int y = page.Bottom - h - Math.Max(6, page.Height / 120);
                var r = new Rectangle(page.Left + Math.Max(12, page.Width / 70), y, page.Width - Math.Max(24, page.Width / 35), h);
                g.DrawString("JIQP-0202-02", footerFont, b, r, left);
                g.DrawString("( \uC8FC ) \uC9C0   \uC778", footerFont, b, r, center);
                g.DrawString("A4(210*297)\u339C", footerFont, b, r, right);
            }
        }
        void DrawAnswerPage(Graphics g, Rectangle bounds)
        {
            if (Doc == null) return;
            MainForm.DrawAnswerTable(g, bounds, Doc.Bank.Name, Doc.UserName, Doc.EvalDate, Doc.Questions, true);
        }
        static void CenterText(Graphics g, string text, Font font, Brush brush, Rectangle rect)
        {
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(text ?? "", font, brush, rect, sf);
        }
        void DrawHeader(Graphics g, Pen pen, Font title, Font font, Font bold, Func<float, int> X, Func<float, int> Y, Func<float, int> W, Func<float, int> H)
        {
            Bank b = Doc.Bank;
            const int dy = 18;
            Cell(g, "\uBD80\uC11C\uBA85", bold, new Rectangle(X(16), Y(8 + dy), W(115), H(24)), true);
            Cell(g, b.Department, font, new Rectangle(X(131), Y(8 + dy), W(130), H(24)), true);
            TextOnly(g, "\uAD50\uC721 \uD3C9\uAC00\uC11C", title, new Rectangle(X(261), Y(8 + dy), W(515), H(72)), true);
            Cell(g, "\uAE30 \uC548", bold, new Rectangle(X(776), Y(8 + dy), W(110), H(24)), true);
            Cell(g, "\uC2EC\uC758", bold, new Rectangle(X(886), Y(8 + dy), W(110), H(24)), true);
            Cell(g, "\uACB0\uC815", bold, new Rectangle(X(996), Y(8 + dy), W(110), H(24)), true);
            Cell(g, "\uD3C9\uAC00\uC790", bold, new Rectangle(X(16), Y(32 + dy), W(115), H(48)), true);
            Cell(g, b.Evaluator, font, new Rectangle(X(131), Y(32 + dy), W(130), H(48)), true);
            Cell(g, "", font, new Rectangle(X(776), Y(32 + dy), W(110), H(48)), true);
            Cell(g, "", font, new Rectangle(X(886), Y(32 + dy), W(110), H(48)), true);
            Cell(g, "", font, new Rectangle(X(996), Y(32 + dy), W(110), H(48)), true);
            Cell(g, "\uD3C9\uAC00 \uC77C\uC2DC", bold, new Rectangle(X(16), Y(80 + dy), W(115), H(24)), true);
            Cell(g, Doc.EvalDate, font, new Rectangle(X(131), Y(80 + dy), W(130), H(24)), true);
            Cell(g, "/", font, new Rectangle(X(776), Y(80 + dy), W(110), H(24)), true);
            Cell(g, "/", font, new Rectangle(X(886), Y(80 + dy), W(110), H(24)), true);
            Cell(g, "/", font, new Rectangle(X(996), Y(80 + dy), W(110), H(24)), true);
            Cell(g, "\uC9C1 \uBB34 \uBA85", bold, new Rectangle(X(16), Y(112 + dy), W(230), H(26)), true);
            Cell(g, b.JobName, font, new Rectangle(X(246), Y(112 + dy), W(470), H(26)), true);
            Cell(g, "\uC131 \uBA85", bold, new Rectangle(X(716), Y(112 + dy), W(90), H(26)), true);
            Cell(g, Doc.UserName, font, new Rectangle(X(806), Y(112 + dy), W(300), H(26)), true);
            Cell(g, "\uD3C9\uAC00 \uBC29\uBC95", bold, new Rectangle(X(16), Y(138 + dy), W(230), H(26)), true);
            Cell(g, "\uC2DC\uD5D8 \uD3C9\uAC00, \uACB0\uACFC \uBCF4\uACE0\uC11C, \uC9C1\uBB34 \uD3C9\uAC00, \uAE30\uD0C0 \uBC29\uBC95 (          )", font, new Rectangle(X(246), Y(138 + dy), W(860), H(26)), true);
            Cell(g, "\uAC1C\uC815 \uCC28\uC218", bold, new Rectangle(X(16), Y(164 + dy), W(230), H(26)), true);
            Cell(g, b.Revision, font, new Rectangle(X(246), Y(164 + dy), W(160), H(26)), true);
            Cell(g, "\uC81C\uC815\uC77C", bold, new Rectangle(X(406), Y(164 + dy), W(160), H(26)), true);
            Cell(g, b.IssueDate, font, new Rectangle(X(566), Y(164 + dy), W(170), H(26)), true);
            Cell(g, "\uC720 \uD615", bold, new Rectangle(X(736), Y(164 + dy), W(150), H(26)), true);
            Cell(g, b.ProductType, font, new Rectangle(X(886), Y(164 + dy), W(220), H(26)), true);
            DrawHeaderGrid(g, pen, X, Y);
        }
        static void Cell(Graphics g, string value, Font font, Rectangle rect, bool center)
        {
            TextOnly(g, value, font, rect, center);
        }
        static void DrawHeaderGrid(Graphics g, Pen pen, Func<float, int> X, Func<float, int> Y)
        {
            const int dy = 18;
            Action<float, float, float> h = (y, x1, x2) => g.DrawLine(pen, X(x1), Y(y), X(x2), Y(y));
            Action<float, float, float> v = (x, y1, y2) => g.DrawLine(pen, X(x), Y(y1), X(x), Y(y2));

            h(8 + dy, 16, 261); h(32 + dy, 16, 261); h(80 + dy, 16, 261); h(104 + dy, 16, 261);
            v(16, 8 + dy, 104 + dy); v(131, 8 + dy, 104 + dy); v(261, 8 + dy, 104 + dy);

            h(8 + dy, 776, 1106); h(32 + dy, 776, 1106); h(80 + dy, 776, 1106); h(104 + dy, 776, 1106);
            v(776, 8 + dy, 104 + dy); v(886, 8 + dy, 104 + dy); v(996, 8 + dy, 104 + dy); v(1106, 8 + dy, 104 + dy);

            h(112 + dy, 16, 1106); h(138 + dy, 16, 1106); h(164 + dy, 16, 1106); h(190 + dy, 16, 1106);
            v(16, 112 + dy, 190 + dy); v(1106, 112 + dy, 190 + dy);
            v(246, 112 + dy, 190 + dy); v(716, 112 + dy, 138 + dy); v(806, 112 + dy, 138 + dy);
            v(406, 164 + dy, 190 + dy); v(566, 164 + dy, 190 + dy); v(736, 164 + dy, 190 + dy); v(886, 164 + dy, 190 + dy);
        }
        static void TextOnly(Graphics g, string value, Font font, Rectangle rect, bool center)
        {
            using (var b = new SolidBrush(Color.Black))
            {
                var sf = new StringFormat { Alignment = center ? StringAlignment.Center : StringAlignment.Near, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                g.DrawString(value ?? "", font, b, new Rectangle(rect.X + 4, rect.Y + 1, rect.Width - 8, rect.Height - 2), sf);
            }
        }
        static void DrawNumber(Graphics g, string value, Font font, Rectangle rect)
        {
            using (var b = new SolidBrush(Color.Black))
            {
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Near };
                g.DrawString(value ?? "", font, b, rect, sf);
            }
        }
        static void QuestionCell2(Graphics g, Question q, int number, Font font, Font bold, Brush brush, Rectangle rect, bool answers)
        {
            string text = number.ToString() + ".  " + FormatQuestionText(q);
            Rectangle textRect = new Rectangle(rect.X + 4, rect.Y + 6, rect.Width - 8, rect.Height - 12);
            if (q.Images.Count > 0)
            {
                int count = Math.Min(4, q.Images.Count);
                string imageText = CompactVisibleLines(StripImageChoiceLines(text));
                string choiceText = CompactVisibleLines(ImageChoiceLines(text));
                List<string> choices = ImageChoiceLineList(choiceText, count);
                bool choicesBelongToImages = ChoicesBelongToImages(count, choices.Count);
                bool wideMultiImage = HasWideMultiImages(q);
                int mainLines = Math.Max(1, VisibleWrappedLineCount(imageText, ScoreWrapUnits));
                int maxImgH = count == 1 ? 145 : 125;
                int gap = count == 1 ? 0 : 18;
                int imgY = rect.Y + Math.Max(38, mainLines * font.Height + 12);
                int availableW = Math.Max(100, rect.Width - 84);
                int imgW = count == 1
                    ? Math.Min(340, availableW)
                    : Math.Min(150, Math.Max(80, (availableW - gap * (count - 1)) / count));
                int startX = rect.X + 42;
                int usedImgH = 0;
                var imageSlots = new List<Rectangle>();
                for (int idx = 0; idx < count; idx++)
                {
                    Image img = q.Images[idx];
                    int minimumW = count == 1 ? 180 : wideMultiImage ? 145 : 128;
                    int iw = Math.Min(imgW, Math.Max(minimumW, img.Width));
                    int ih = img.Height * iw / Math.Max(1, img.Width);
                    if (ih > maxImgH)
                    {
                        ih = maxImgH;
                        iw = img.Width * ih / Math.Max(1, img.Height);
                    }
                    int x = startX + idx * (imgW + gap) + (imgW - iw) / 2;
                    g.DrawImage(img, x, imgY, iw, ih);
                    if (count > 1)
                    {
                        string marker = ChoiceMarker(idx);
                        using (var markerFont = new Font(font.FontFamily, font.Size + 1.5f, FontStyle.Regular))
                        using (var markerFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            int markerW = Math.Max(20, markerFont.Height + 4);
                            if (!wideMultiImage)
                                g.FillRectangle(Brushes.White, new Rectangle(x - 2, imgY - 1, markerW, markerFont.Height + 5));
                            int markerX = Math.Max(rect.X + 4, x - markerW - 3);
                            var markerRect = new Rectangle(markerX, imgY - 1, markerW, markerFont.Height + 5);
                            g.FillRectangle(Brushes.White, markerRect);
                            g.DrawString(marker, markerFont, brush, markerRect, markerFormat);
                        }
                    }
                    imageSlots.Add(new Rectangle(startX + idx * (imgW + gap), imgY, imgW, ih));
                    usedImgH = Math.Max(usedImgH, ih);
                }
                textRect.Height = Math.Max(22, imgY - textRect.Y - 4);
                DrawSpacedText(g, imageText, font, brush, textRect, 0, ScoreWrapUnits);

                int captionHeight = 0;
                if (choicesBelongToImages && choices.Count > 0)
                {
                    bool hasCaption = false;
                    using (var center = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Near })
                    {
                        for (int idx = 0; idx < Math.Min(count, choices.Count); idx++)
                        {
                            string caption = ImageChoiceCaption(choices[idx]);
                            if (caption.Length == 0) continue;
                            hasCaption = true;
                            var captionRect = new Rectangle(imageSlots[idx].X, imgY + usedImgH + 4, imageSlots[idx].Width, font.Height + 4);
                            g.DrawString(caption, font, brush, captionRect, center);
                        }
                    }
                    if (hasCaption) captionHeight = font.Height + 6;
                }

                string remainingChoices = choicesBelongToImages
                    ? string.Join("\n", choices.Skip(Math.Min(count, choices.Count)))
                    : choiceText;
                if (remainingChoices.Trim().Length > 0)
                {
                    var choiceRect = new Rectangle(rect.X + 42, imgY + usedImgH + captionHeight + 8, rect.Width - 80, Math.Max(16, rect.Bottom - imgY - usedImgH - captionHeight - 10));
                    DrawSpacedText(g, remainingChoices, font, brush, choiceRect, 0);
                }
                DrawQuestionScore(g, q, font, brush, rect);
                if (answers)
                {
                    using (var red = new SolidBrush(Color.FromArgb(200, 0, 0)))
                        g.DrawString(AnswerMark(q), bold, red, rect.Right - 128, rect.Y + 5);
                }
                return;
            }
            DrawSpacedText(g, text, font, brush, textRect, q.Type == "\uC8FC\uAD00\uC2DD" ? 2 : 0, ScoreWrapUnits);
            DrawQuestionScore(g, q, font, brush, rect);
            if (answers)
            {
                using (var red = new SolidBrush(Color.FromArgb(200, 0, 0)))
                    g.DrawString(AnswerMark(q), bold, red, rect.Right - 128, rect.Y + 5);
            }
        }
        static string StripImageChoiceLines(string text)
        {
            var kept = new List<string>();
            foreach (string line in (text ?? "").Split('\n'))
            {
                if (IsImageChoiceLine(line))
                    continue;
                kept.Add(line);
            }
            return string.Join("\n", kept);
        }
        static string ImageChoiceLines(string text)
        {
            var kept = new List<string>();
            foreach (string line in (text ?? "").Split('\n'))
                if (IsImageChoiceLine(line))
                    kept.Add(line);
            return string.Join("\n", kept);
        }
        static List<string> ImageChoiceLineList(string text, int imageCount)
        {
            var result = new List<string>();
            foreach (string line in (text ?? "").Split('\n'))
            {
                string value = line.Trim();
                if (value.Length == 0) continue;
                var markers = Regex.Matches(value, "[\u2460\u2461\u2462\u2463\u2464\u2465]");
                if (imageCount > 1 && markers.Count > 1)
                {
                    for (int i = 0; i < markers.Count; i++)
                    {
                        int start = markers[i].Index;
                        int end = i + 1 < markers.Count ? markers[i + 1].Index : value.Length;
                        result.Add(value.Substring(start, end - start).Trim());
                    }
                }
                else
                {
                    result.Add(value);
                }
            }
            return result;
        }
        static bool ChoicesBelongToImages(int imageCount, int choiceCount)
        {
            int count = Math.Min(4, imageCount);
            return count > 1 && (choiceCount == 0 || choiceCount >= count);
        }
        static string ImageChoiceCaption(string line)
        {
            return Regex.Replace((line ?? "").Trim(), @"^[\u2460\u2461\u2462\u2463\u2464\u2465]\s*", "").Trim();
        }
        static string ChoiceMarker(int index)
        {
            const string markers = "\u2460\u2461\u2462\u2463\u2464\u2465";
            return index >= 0 && index < markers.Length ? markers[index].ToString() : (index + 1).ToString();
        }
        static bool HasWideMultiImages(Question q)
        {
            int count = Math.Min(4, q.Images.Count);
            if (count <= 1) return false;
            double averageRatio = q.Images.Take(count).Average(img => img.Width / (double)Math.Max(1, img.Height));
            return averageRatio >= 1.1;
        }
        static bool IsImageChoiceLine(string line)
        {
            string value = (line ?? "").Trim();
            return value.Length > 0 && Regex.IsMatch(value, @"^[①②③④⑤⑥]");
        }
        static string CompactVisibleLines(string text)
        {
            return string.Join("\n", (text ?? "").Split('\n').Where(line => line.Trim().Length > 0));
        }
        static int VisibleWrappedLineCount(string text, int firstLineUnits = PrintWrapUnits)
        {
            return WrapPrintLines(text, firstLineUnits).Count(line => line.Trim().Length > 0);
        }
        static List<string> WrapPrintLines(string text)
        {
            return WrapPrintLines(text, PrintWrapUnits);
        }
        static List<string> WrapPrintLines(string text, int firstLineUnits)
        {
            var result = new List<string>();
            bool firstVisibleLine = true;
            foreach (string sourceLine in (text ?? "").Split('\n'))
            {
                if (sourceLine.Trim().Length == 0)
                {
                    result.Add("");
                    continue;
                }
                string remaining = sourceLine.TrimEnd();
                string continuation = ContinuationIndent(sourceLine);
                int lineLimit = firstVisibleLine ? Math.Max(20, firstLineUnits) : PrintWrapUnits;
                while (VisualUnits(remaining) > lineLimit)
                {
                    int cut = FindWrapIndex(remaining, lineLimit);
                    result.Add(remaining.Substring(0, cut).TrimEnd());
                    firstVisibleLine = false;
                    remaining = remaining.Substring(cut).TrimStart();
                    if (remaining.Length > 0)
                        remaining = continuation + remaining;
                    lineLimit = PrintWrapUnits;
                }
                if (remaining.Length > 0)
                {
                    result.Add(remaining);
                    firstVisibleLine = false;
                }
            }
            return result;
        }
        static string ContinuationIndent(string line)
        {
            string value = (line ?? "").TrimStart();
            if (Regex.IsMatch(value, @"^\d+\.\s*")) return "     ";
            if (Regex.IsMatch(value, @"^[①②③④⑤⑥]")) return "   ";
            return "     ";
        }
        static int FindWrapIndex(string value, int maxUnits)
        {
            int units = 0;
            int lastSpace = -1;
            for (int i = 0; i < value.Length; i++)
            {
                int next = CharUnits(value[i]);
                if (units + next > maxUnits)
                {
                    if (lastSpace >= Math.Max(1, i / 2)) return lastSpace + 1;
                    return Math.Max(1, i);
                }
                units += next;
                if (char.IsWhiteSpace(value[i])) lastSpace = i;
            }
            return value.Length;
        }
        static int VisualUnits(string value)
        {
            int units = 0;
            foreach (char ch in value ?? "") units += CharUnits(ch);
            return units;
        }
        static int CharUnits(char ch)
        {
            if (ch == '\t') return 4;
            return ch <= 0x7f ? 1 : 2;
        }
        static void DrawSpacedText(Graphics g, string text, Font font, Brush brush, Rectangle rect, int extra, int firstLineUnits = PrintWrapUnits)
        {
            int y = rect.Y;
            foreach (string line in WrapPrintLines(text, firstLineUnits))
            {
                if (line.Trim().Length == 0)
                {
                    y += font.Height + extra;
                    continue;
                }
                g.DrawString(line, font, brush, rect.X, y);
                y += font.Height + extra;
                if (y > rect.Bottom) break;
            }
        }
        static void DrawQuestionScore(Graphics g, Question q, Font font, Brush brush, Rectangle rect)
        {
            var format = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Near };
            var scoreRect = new Rectangle(rect.Right - 68, rect.Y + 5, 64, font.Height + 4);
            g.DrawString(ScoreText(q), font, brush, scoreRect, format);
        }
        static void QuestionCell(Graphics g, Question q, Font font, Font bold, Brush brush, Rectangle rect, bool answers)
        {
            string text = FormatQuestionText(q);
            Rectangle textRect = new Rectangle(rect.X + 8, rect.Y + 6, rect.Width - 16, rect.Height - 12);
            if (q.Images.Count > 0)
            {
                int count = Math.Min(4, q.Images.Count);
                int imgH = 82;
                int gap = 34;
                int imgY = rect.Bottom - imgH - 10;
                int imgW = 72;
                int totalW = count * imgW + (count - 1) * gap;
                int startX = rect.X + Math.Max(18, (rect.Width - totalW) / 2);
                for (int idx = 0; idx < count; idx++)
                {
                    Image img = q.Images[idx];
                    int iw = imgW;
                    int ih = img.Height * iw / Math.Max(1, img.Width);
                    if (ih > imgH)
                    {
                        ih = imgH;
                        iw = img.Width * ih / Math.Max(1, img.Height);
                    }
                    int x = startX + idx * (imgW + gap) + (imgW - iw) / 2;
                    g.DrawString(new[] { "\u2460", "\u2461", "\u2462", "\u2463" }[idx], font, brush, x - 18, imgY + imgH / 2 - 7);
                    g.DrawImage(img, x, imgY, iw, ih);
                }
                textRect.Height = Math.Max(24, imgY - textRect.Y - 6);
            }
            g.DrawString(text, font, brush, textRect);
            if (q.Type == "\uC8FC\uAD00\uC2DD")
            {
                int y = rect.Bottom - 26;
                g.DrawLine(Pens.Black, rect.X + 20, y, rect.Right - 20, y);
            }
            if (answers)
            {
                using (var red = new SolidBrush(Color.FromArgb(200, 0, 0)))
                    g.DrawString(AnswerMark(q), bold, red, rect.Right - 70, rect.Y + 5);
            }
        }
        static string FormatQuestionText(Question q)
        {
            string text = q.Text ?? "";
            if (q.Type == "\uC8FC\uAD00\uC2DD")
                text = ExpandSubjectiveBlanks(text);
            string[] lines = text.Split('\n');
            if (lines.Length > 0)
            {
                string scorePattern = @"\s*\(\s*\d+(?:\.\d+)?\s*\uC810\s*\)";
                lines[0] = Regex.Replace(lines[0], scorePattern, "");
            }
            return string.Join("\n", lines);
        }
        static string ScoreText(Question q)
        {
            return "(" + q.Score.ToString("0.##") + "\uC810)";
        }
        static string ExpandSubjectiveBlanks(string text)
        {
            string[] lines = (text ?? "").Split('\n');
            bool hasAnswerBlankAfterFirstLine = lines.Skip(1).Any(line => Regex.IsMatch(line, @"\([ \t\u3000]*\)"));
            int firstExpandableLine = hasAnswerBlankAfterFirstLine ? 1 : 0;
            for (int i = firstExpandableLine; i < lines.Length; i++)
            {
                lines[i] = Regex.Replace(lines[i], @"\([ \t\u3000]*\)", delegate(Match m)
                {
                    int current = Math.Max(0, m.Value.Length - 2);
                    return "(" + new string(' ', Math.Max(20, current)) + ")";
                });
            }
            return string.Join("\n", lines);
        }
    }

    internal sealed class PreviewWheelMessageFilter : IMessageFilter, IDisposable
    {
        const int WmMouseWheel = 0x020A;
        readonly Form dialog;
        readonly PrintPreviewControl preview;
        readonly int pageCount;
        bool disposed;

        public PreviewWheelMessageFilter(Form dialog, PrintPreviewControl preview, int pageCount)
        {
            this.dialog = dialog;
            this.preview = preview;
            this.pageCount = Math.Max(1, pageCount);
            Application.AddMessageFilter(this);
        }

        public bool PreFilterMessage(ref Message message)
        {
            if (message.Msg != WmMouseWheel || disposed || dialog.IsDisposed || !dialog.Visible || pageCount <= 1)
                return false;
            if (!dialog.RectangleToScreen(dialog.ClientRectangle).Contains(Cursor.Position))
                return false;

            int delta = unchecked((short)(((long)message.WParam >> 16) & 0xffff));
            if (delta == 0)
                return false;

            int next = preview.StartPage + (delta < 0 ? 1 : -1);
            preview.StartPage = Math.Max(0, Math.Min(pageCount - 1, next));
            preview.Invalidate();
            return true;
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Application.RemoveMessageFilter(this);
        }
    }

    internal sealed class MainForm : Form
    {
#if UI04
        static readonly int UiVariant = 4;
#elif UI18
        static readonly int UiVariant = 18;
#elif UI31
        static readonly int UiVariant = 31;
#else
        static readonly int UiVariant = 1;
#endif
        readonly Reader reader = new Reader();
        readonly Random random = new Random();
        readonly List<Bank> banks = new List<Bank>();
        readonly List<Bank> visibleBanks = new List<Bank>();
        ExamSettings examSettings = ExamSettingsStore.Load();
        TextBox pathBox;
        TextBox userNameBox, evalDateBox;
        DataGridView grid;
        NumericUpDown vdaCount, choiceCount, subjectiveCount, vdaScore, choiceScore, subjectiveScore, targetScore;
        Control vdaRow;
        Label totalLabel, pageLabel, status;
        CheckBox answerCheck;
        Preview preview;
        ExamDoc doc;
        Panel landingPage, workspacePage;
        Label audienceLabel, selectedBankNameLabel, selectedBankMetaLabel, generationLabel, stepLabel, secondaryStatus;
        string audienceFilter = "";

        public MainForm()
        {
            Text = "OJT EXAM MAKER";
            Width = 1760; Height = 900; MinimumSize = new Size(1280, 720);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9);
            BackColor = Color.FromArgb(11, 18, 32);
            BuildUi();
            pathBox.Text = DefaultWorkbook();
            LoadBanks();
        }
        void BuildLegacyUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18, 28, 14, 14), RowCount = 3, ColumnCount = 3 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 540));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            Controls.Add(root);
            var title = new Label { Text = "OJT EXAM MAKER", ForeColor = Color.White, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 20, FontStyle.Bold), Padding = new Padding(0, 8, 0, 0) };
            root.Controls.Add(title, 0, 0); root.SetColumnSpan(title, 3);
            status = new Label { Text = "", Visible = false, Width = 1, Height = 1 };
            root.Controls.Add(status, 2, 0);
            var file = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, BackColor = Color.FromArgb(18, 28, 46), Padding = new Padding(8) };
            file.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
            file.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            file.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            root.Controls.Add(file, 0, 1); root.SetColumnSpan(file, 3);
            file.Controls.Add(Label("\uBB38\uC81C\uC740\uD589"), 0, 0);
            pathBox = new TextBox { Dock = DockStyle.Fill, BackColor = Color.FromArgb(8, 17, 31), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            file.Controls.Add(pathBox, 1, 0);
            file.Controls.Add(Button("\uD30C\uC77C \uC120\uD0DD", Browse), 2, 0);
            root.Controls.Add(BankPanel(), 0, 2);
            root.Controls.Add(SettingsPanel(), 1, 2);
            root.Controls.Add(PreviewPanel(), 2, 2);
        }
        Control BankPanel()
        {
            var p = Panel("\uACF5\uC815 \uC120\uD0DD");
            grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeColumns = false, AllowUserToResizeRows = false, RowHeadersVisible = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None, BackgroundColor = Color.FromArgb(18, 28, 46), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 45, 72); grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.DefaultCellStyle.BackColor = Color.FromArgb(18, 28, 46); grid.DefaultCellStyle.ForeColor = Color.White; grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(20, 81, 63); grid.DefaultCellStyle.SelectionForeColor = Color.White; grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.Columns.Add("name", "\uACF5\uC815\uBA85"); grid.Columns.Add("vda", "VDA"); grid.Columns.Add("choice", "\uAC1D\uAD00\uC2DD"); grid.Columns.Add("subjective", "\uC8FC\uAD00\uC2DD");
            grid.Columns[0].FillWeight = 280; grid.Columns[1].FillWeight = 58; grid.Columns[2].FillWeight = 70; grid.Columns[3].FillWeight = 70;
            grid.Columns[0].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.RowTemplate.Height = 28;
            grid.ColumnHeadersHeight = 26;
            grid.CellFormatting += delegate(object s, DataGridViewCellFormattingEventArgs e) { if (e.RowIndex >= 0 && e.ColumnIndex == 0 && e.Value != null) grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = e.Value.ToString(); };
            grid.SelectionChanged += delegate { RefreshSelected(); };
            p.Controls.Add(grid); return p;
        }
        Control SettingsPanel()
        {
            var p = Panel("\uC2DC\uD5D8 \uC870\uAC74");
            var body = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(14), AutoScroll = true };
            p.Controls.Add(body);
            var conditionHeader = new TableLayoutPanel { Width = 470, Height = 30, ColumnCount = 3, Margin = new Padding(0, 0, 0, 2) };
            conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f)); conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f));
            var settingsButton = Button("\uC124\uC815", OpenExamSettings); settingsButton.Dock = DockStyle.Top; settingsButton.Height = 24; settingsButton.Margin = new Padding(0); conditionHeader.Controls.Add(settingsButton, 0, 0);
            conditionHeader.Controls.Add(new Label { Text = "\uBB38\uC81C\uC218", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold), TextAlign = ContentAlignment.BottomCenter, Padding = new Padding(0, 0, 0, 1) }, 1, 0);
            conditionHeader.Controls.Add(new Label { Text = "\uC810\uC218", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold), TextAlign = ContentAlignment.BottomCenter, Padding = new Padding(0, 0, 0, 1) }, 2, 0);
            body.Controls.Add(conditionHeader);
            vdaRow = Row("VDA", out vdaCount, out vdaScore, 2, 2.5m, true); body.Controls.Add(vdaRow);
            body.Controls.Add(Row("\uAC1D\uAD00\uC2DD", out choiceCount, out choiceScore, 20, 4m, true));
            body.Controls.Add(Row("\uC8FC\uAD00\uC2DD", out subjectiveCount, out subjectiveScore, 4, 5m, true));
            targetScore = Num(0, 1000, 100, 0); body.Controls.Add(TargetRow());
            totalLabel = new Label { Text = "TOTAL 100", Width = 470, Height = 52, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(32, 208, 132), BackColor = Color.FromArgb(16, 36, 58), Font = new Font("Segoe UI", 25, FontStyle.Bold), Margin = new Padding(0, 8, 0, 0) };
            body.Controls.Add(totalLabel);
            body.Controls.Add(InputRow("\uC131\uBA85", out userNameBox, ""));
            body.Controls.Add(InputRow("\uD3C9\uAC00 \uC77C\uC2DC", out evalDateBox, DateTime.Today.ToString("yyyy.MM.dd")));
            answerCheck = new CheckBox { Text = "\uBBF8\uB9AC\uBCF4\uAE30\uC5D0 \uB2F5\uC548 \uD45C\uC2DC", Width = 470, ForeColor = Color.White, Margin = new Padding(0, 10, 0, 0) };
            answerCheck.CheckedChanged += delegate { if (doc != null) { doc.ShowAnswers = answerCheck.Checked; preview.Invalidate(); } };
            body.Controls.Add(answerCheck);
            var nav = new TableLayoutPanel { Width = 470, Height = 36, ColumnCount = 3, Margin = new Padding(0, 6, 0, 0) };
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            nav.Controls.Add(Button("\uC774\uC804", delegate { preview.MovePage(-1); UpdatePage(); }), 0, 0);
            pageLabel = new Label { Text = "1 / 1", Dock = DockStyle.Fill, ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 10, FontStyle.Bold) }; nav.Controls.Add(pageLabel, 1, 0);
            nav.Controls.Add(Button("\uB2E4\uC74C", delegate { preview.MovePage(1); UpdatePage(); }), 2, 0);
            body.Controls.Add(nav);
            body.Controls.Add(Action("\uB79C\uB364 \uBBF8\uB9AC\uBCF4\uAE30", Generate, Color.FromArgb(32, 208, 132)));
            body.Controls.Add(Action("\uCD9C\uB825 \uBBF8\uB9AC\uBCF4\uAE30", delegate { PrintPreviewProblemAndAnswer(); }, Color.FromArgb(30, 45, 72)));
            body.Controls.Add(Action("\uBB38\uC81C + \uB2F5\uC548 \uD504\uB9B0\uD2B8", delegate { PrintProblemAndAnswer(); }, Color.FromArgb(30, 45, 72)));
            return p;
        }
        Control PreviewPanel()
        {
            var p = Panel("\uC2DC\uD5D8\uC9C0 \uBBF8\uB9AC\uBCF4\uAE30");
            preview = new Preview { Dock = DockStyle.Fill };
            p.Controls.Add(preview);
            var zoom = new FlowLayoutPanel { Width = 150, Height = 30, Left = p.Width - 168, Top = 12, Anchor = AnchorStyles.Top | AnchorStyles.Right, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, BackColor = Color.FromArgb(18, 28, 46) };
            zoom.Controls.Add(SmallButton("-", delegate { preview.ZoomBy(1f / 1.12f); }));
            zoom.Controls.Add(SmallButton("100%", delegate { preview.Fit(); }));
            zoom.Controls.Add(SmallButton("+", delegate { preview.ZoomBy(1.12f); }));
            p.Controls.Add(zoom);
            zoom.BringToFront();
            return p;
        }
        Button SmallButton(string text, EventHandler h) { var b = Button(text, h); b.Width = 46; b.Height = 26; b.Dock = DockStyle.None; b.Margin = new Padding(2, 0, 0, 0); return b; }
        Panel Panel(string title) { var p = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 28, 46), Padding = new Padding(14, 48, 14, 14), Margin = new Padding(0, 0, 12, 0) }; var l = new Label { Text = title, Left = 14, Top = 12, Height = 30, Width = 500, ForeColor = Color.White, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 14, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right }; p.Controls.Add(l); l.BringToFront(); return p; }
        Label Label(string text) { return new Label { Text = text, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold) }; }
        Button Button(string text, EventHandler h) { var b = new Button { Text = text, Dock = DockStyle.Fill, BackColor = Color.FromArgb(30, 45, 72), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold), Margin = new Padding(4, 0, 0, 0) }; b.Click += h; return b; }
        Button Action(string text, EventHandler h, Color c) { var b = new Button { Text = text, Width = 470, Height = 44, BackColor = c, ForeColor = c == Color.FromArgb(32, 208, 132) ? Color.Black : Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 10, FontStyle.Bold), Margin = new Padding(0, 10, 0, 0) }; b.Click += h; return b; }
        Control Row(string name, out NumericUpDown count, out NumericUpDown score, int cv, decimal sv, bool wide)
        {
            var row = new TableLayoutPanel { Width = 470, Height = 44, ColumnCount = 3 };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f));
            row.Controls.Add(new Label { Text = name, Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 12, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            count = Num(0, 500, cv, 0); score = Num(0, 100, sv, 1); row.Controls.Add(count, 1, 0); row.Controls.Add(score, 2, 0); return row;
        }
        Control TargetRow() { var row = new TableLayoutPanel { Width = 470, Height = 44, ColumnCount = 2 }; row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28)); row.Controls.Add(new Label { Text = "\uBAA9\uD45C \uC810\uC218", Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0); row.Controls.Add(targetScore, 1, 0); return row; }
        Control InputRow(string name, out TextBox box, string value)
        {
            var row = new TableLayoutPanel { Width = 470, Height = 44, ColumnCount = 2, Margin = new Padding(0, 5, 0, 0) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            row.Controls.Add(new Label { Text = name, Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            box = new TextBox { Text = value, Dock = DockStyle.Fill, BackColor = Color.FromArgb(8, 17, 31), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11, FontStyle.Bold), TextAlign = HorizontalAlignment.Center, Margin = new Padding(0, 4, 0, 4) };
            box.TextChanged += delegate { if (doc != null) { doc.UserName = userNameBox.Text.Trim(); doc.EvalDate = evalDateBox.Text.Trim(); preview.Invalidate(); } };
            row.Controls.Add(box, 1, 0);
            return row;
        }
        NumericUpDown Num(decimal min, decimal max, decimal value, int dec)
        {
            var n = new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = value,
                DecimalPlaces = dec,
                Increment = 0m,
                ReadOnly = true,
                InterceptArrowKeys = false,
                TabStop = false,
                TextAlign = HorizontalAlignment.Center,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(8, 17, 31),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            if (n.Controls.Count > 0) n.Controls[0].Enabled = false;
            n.ValueChanged += delegate { UpdateTotal(); };
            return n;
        }

        sealed class UiPalette
        {
            public Color Background;
            public Color Surface;
            public Color SurfaceAlt;
            public Color Foreground;
            public Color Muted;
            public Color Border;
            public Color Accent;
            public Color AccentText;
            public Color Header;
            public Color Input;
        }

        void BuildUi()
        {
            Controls.Clear();
            Width = 1600;
            Height = 900;
            MinimumSize = new Size(1280, 720);
            Text = "OJT EXAM MAKER";

            if (UiVariant == 4) BuildUi04();
            else if (UiVariant == 18) BuildUi18();
            else if (UiVariant == 31) BuildUi31();
            else BuildUi01();
        }

        string VariantName()
        {
            if (UiVariant == 4) return "04 대화형 안내";
            if (UiVariant == 18) return "18 스마트 팩토리 맵";
            if (UiVariant == 31) return "31 Carbon 데이터 작업대";
            return "01 네 개의 시작 문";
        }

        UiPalette LightPalette(Color accent)
        {
            return new UiPalette
            {
                Background = Color.FromArgb(238, 244, 242),
                Surface = Color.White,
                SurfaceAlt = Color.FromArgb(226, 235, 232),
                Foreground = Color.FromArgb(17, 31, 34),
                Muted = Color.FromArgb(83, 107, 108),
                Border = Color.FromArgb(166, 184, 181),
                Accent = accent,
                AccentText = Color.White,
                Header = Color.FromArgb(21, 39, 43),
                Input = Color.White
            };
        }

        UiPalette DarkPalette(Color accent)
        {
            return new UiPalette
            {
                Background = Color.FromArgb(20, 34, 39),
                Surface = Color.FromArgb(33, 51, 58),
                SurfaceAlt = Color.FromArgb(25, 42, 48),
                Foreground = Color.FromArgb(244, 248, 247),
                Muted = Color.FromArgb(175, 195, 199),
                Border = Color.FromArgb(75, 99, 107),
                Accent = accent,
                AccentText = Color.White,
                Header = Color.FromArgb(17, 29, 33),
                Input = Color.FromArgb(18, 31, 36)
            };
        }

        void BuildUi01()
        {
            UiPalette p = LightPalette(Color.FromArgb(0, 142, 122));
            BackColor = p.Background;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = p.Background, Padding = new Padding(18) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);
            root.Controls.Add(BuildHeader(p, "OJT EXAM MAKER", "QUALITY TRAINING · PRINT SYSTEM"), 0, 0);
            root.Controls.Add(BuildFileBar(p), 0, 1);
            var host = new Panel { Dock = DockStyle.Fill, BackColor = p.Background };
            root.Controls.Add(host, 0, 2);
            landingPage = BuildAudiencePage(p, 1);
            workspacePage = BuildThreeColumnWorkspace(p, true);
            workspacePage.Visible = false;
            host.Controls.Add(workspacePage);
            host.Controls.Add(landingPage);
            landingPage.BringToFront();
        }

        void BuildUi04()
        {
            UiPalette p = LightPalette(Color.FromArgb(45, 121, 77));
            BackColor = p.Background;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 1, ColumnCount = 2, BackColor = p.Background };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 278));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(root);
            root.Controls.Add(BuildGuideSidebar(p), 0, 0);
            var host = new Panel { Dock = DockStyle.Fill, Padding = new Padding(28), BackColor = p.Background };
            root.Controls.Add(host, 1, 0);
            landingPage = BuildAudiencePage(p, 4);
            workspacePage = BuildThreeColumnWorkspace(p, true);
            workspacePage.Visible = false;
            host.Controls.Add(workspacePage);
            host.Controls.Add(landingPage);
            landingPage.BringToFront();
        }

        void BuildUi18()
        {
            UiPalette p = DarkPalette(Color.FromArgb(45, 103, 218));
            BackColor = p.Background;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = p.Background, Padding = new Padding(14) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);
            root.Controls.Add(BuildHeader(p, "OJT FACTORY MAKER", "QUALITY TRAINING · PAPER EXAM"), 0, 0);
            root.Controls.Add(BuildFileBar(p), 0, 1);
            var host = new Panel { Dock = DockStyle.Fill, BackColor = p.Background };
            root.Controls.Add(host, 0, 2);
            landingPage = BuildAudiencePage(p, 18);
            workspacePage = BuildThreeColumnWorkspace(p, true);
            workspacePage.Visible = false;
            host.Controls.Add(workspacePage);
            host.Controls.Add(landingPage);
            landingPage.BringToFront();
        }

        void BuildUi31()
        {
            UiPalette p = LightPalette(Color.FromArgb(0, 137, 123));
            p.Background = Color.FromArgb(245, 247, 247);
            p.Header = Color.FromArgb(20, 20, 20);
            p.Border = Color.FromArgb(178, 185, 186);
            BackColor = p.Background;

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = p.Background };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);
            root.Controls.Add(BuildHeader(p, "OJT DATA WORKSPACE", "QUALITY TRAINING · PRINT CONTROL"), 0, 0);

            var content = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 1, ColumnCount = 2, BackColor = p.Background };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(content, 0, 1);
            content.Controls.Add(BuildCarbonSidebar(p), 0, 0);
            workspacePage = BuildCarbonWorkspace(p);
            content.Controls.Add(workspacePage, 1, 0);
        }

        Control BuildHeader(UiPalette p, string brand, string subtitle)
        {
            var header = new Panel { Dock = DockStyle.Fill, BackColor = p.Header, Padding = new Padding(14, 6, 14, 6), Margin = new Padding(0) };
            var brandPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) };
            brandPanel.Controls.Add(new Label { Text = brand, Left = 42, Top = 2, Width = 360, Height = 26, ForeColor = Color.White, Font = new Font("Segoe UI", 15, FontStyle.Bold) });
            brandPanel.Controls.Add(new Label { Text = subtitle, Left = 43, Top = 28, Width = 360, Height = 18, ForeColor = Color.FromArgb(150, 191, 191), Font = new Font("Segoe UI", 7, FontStyle.Regular) });
            brandPanel.Controls.Add(new Label { Text = "✓", Left = 0, Top = 2, Width = 36, Height = 36, TextAlign = ContentAlignment.MiddleCenter, BackColor = p.Accent, ForeColor = Color.White, Font = new Font("Segoe UI", 16, FontStyle.Bold) });
            status = new Label { Text = "문제은행 확인 중", Dock = DockStyle.Right, Width = 210, TextAlign = ContentAlignment.MiddleRight, ForeColor = Color.FromArgb(188, 211, 211), Font = new Font("맑은 고딕", 9, FontStyle.Bold), Padding = new Padding(0, 0, 12, 0) };
            Control admin = MakeHeaderAction("관리자", OpenExamSettings, p);
            admin.Dock = DockStyle.Right;
            admin.Width = 112;
            header.Controls.Add(brandPanel);
            header.Controls.Add(status);
            header.Controls.Add(admin);
            return header;
        }

        Control BuildFileBar(UiPalette p)
        {
            var bar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, BackColor = p.SurfaceAlt, Padding = new Padding(8), Margin = new Padding(0, 2, 0, 4) };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            bar.Controls.Add(new Label { Text = "문제은행", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = p.Foreground, Font = new Font("맑은 고딕", 9, FontStyle.Bold) }, 0, 0);
            pathBox = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.FixedSingle, BackColor = p.Input, ForeColor = p.Foreground, Font = new Font("맑은 고딕", 9), Margin = new Padding(0, 3, 8, 3) };
            bar.Controls.Add(pathBox, 1, 0);
            bar.Controls.Add(MakeButton("파일 선택", Browse, p.Surface, p.Foreground, p.Border, 9), 2, 0);
            return bar;
        }

        Control BuildGuideSidebar(UiPalette p)
        {
            var side = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 9, ColumnCount = 1, BackColor = Color.FromArgb(20, 39, 43), Padding = new Padding(22) };
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 86));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            side.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            side.Controls.Add(new Label { Text = "✓  OJT EXAM MAKER", Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("Segoe UI", 14, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            stepLabel = new Label { Text = "2 / 4 단계", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(91, 181, 117), BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(8), Font = new Font("맑은 고딕", 9, FontStyle.Bold) };
            side.Controls.Add(stepLabel, 0, 1);
            side.Controls.Add(new Label { Text = "시험지 출력 준비\r\n\r\n대상과 공정을 선택한 뒤\r\n출력 정보를 확인합니다.", Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("맑은 고딕", 11, FontStyle.Bold) }, 0, 2);
            side.Controls.Add(new Label { Text = "문제은행 파일", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(167, 198, 199), Font = new Font("맑은 고딕", 8, FontStyle.Bold), TextAlign = ContentAlignment.BottomLeft }, 0, 3);
            pathBox = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, BackColor = Color.FromArgb(13, 28, 31), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("맑은 고딕", 8), Margin = new Padding(0, 6, 0, 6) };
            side.Controls.Add(pathBox, 0, 4);
            side.Controls.Add(MakeButton("문제은행 선택", Browse, Color.White, p.Foreground, p.Border, 9), 0, 5);
            side.Controls.Add(MakeButton("관리자 설정", OpenExamSettings, Color.FromArgb(38, 67, 70), Color.White, Color.FromArgb(82, 116, 117), 9), 0, 6);
            status = new Label { Text = "문제은행 확인 중", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(126, 205, 153), TextAlign = ContentAlignment.BottomLeft, Font = new Font("맑은 고딕", 8, FontStyle.Bold) };
            side.Controls.Add(status, 0, 7);
            side.Controls.Add(MakeButton("처음부터", delegate { ShowLanding(); }, Color.White, p.Foreground, p.Border, 9), 0, 8);
            return side;
        }

        Control BuildCarbonSidebar(UiPalette p)
        {
            var side = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 10, ColumnCount = 1, BackColor = Color.FromArgb(36, 36, 36), Padding = new Padding(0, 18, 0, 18) };
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            side.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            side.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            side.Controls.Add(NavLabel("▶  시험지 출력", true, p), 0, 0);
            side.Controls.Add(NavLabel("▤  문제은행", false, p), 0, 1);
            side.Controls.Add(NavLabel("▣  출력 문서", false, p), 0, 2);
            side.Controls.Add(NavLabel("◷  출력 이력", false, p), 0, 3);
            side.Controls.Add(NavLabel("⚙  설정", false, p), 0, 4);
            side.Controls.Add(MakeButton("문제은행 선택", Browse, Color.FromArgb(50, 50, 50), Color.White, Color.FromArgb(90, 90, 90), 9), 0, 5);
            secondaryStatus = new Label { Text = "문제은행 확인 중", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(90, 196, 170), Padding = new Padding(18, 0, 12, 0), TextAlign = ContentAlignment.BottomLeft, Font = new Font("맑은 고딕", 8, FontStyle.Bold) };
            side.Controls.Add(secondaryStatus, 0, 6);
            side.Controls.Add(new Label { Text = "연결된 파일", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(150, 150, 150), Padding = new Padding(18, 0, 12, 0), Font = new Font("맑은 고딕", 8, FontStyle.Bold) }, 0, 7);
            pathBox = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, BackColor = Color.FromArgb(28, 28, 28), ForeColor = Color.FromArgb(215, 215, 215), BorderStyle = BorderStyle.None, Font = new Font("맑은 고딕", 8), Margin = new Padding(18, 4, 12, 4) };
            side.Controls.Add(pathBox, 0, 8);
            side.Controls.Add(MakeButton("관리자 설정", OpenExamSettings, Color.FromArgb(50, 50, 50), Color.White, Color.FromArgb(90, 90, 90), 9), 0, 9);
            return side;
        }

        Control NavLabel(string text, bool active, UiPalette p)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                BackColor = active ? p.Accent : Color.FromArgb(36, 36, 36),
                ForeColor = active ? Color.White : Color.FromArgb(155, 178, 181),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(18, 0, 0, 0),
                Font = new Font("맑은 고딕", 9, active ? FontStyle.Bold : FontStyle.Regular)
            };
        }

        Panel BuildAudiencePage(UiPalette p, int style)
        {
            var page = new Panel { Dock = DockStyle.Fill, BackColor = p.Background };
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = p.Background, Padding = style == 4 ? new Padding(14) : new Padding(0, 12, 0, 0) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, style == 18 ? 82 : 112));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            page.Controls.Add(root);

            var heading = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = style == 4 ? p.Surface : p.Background,
                Padding = style == 18 ? new Padding(16, 4, 16, 4) : new Padding(20, 12, 20, 8),
                Margin = new Padding(0),
                RowCount = 2,
                ColumnCount = 2
            };
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, style == 18 ? 42 : 48));
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, style == 18 ? 40 : 48));
            heading.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            heading.Controls.Add(MakeLineIcon("paper", p.Accent, heading.BackColor, false), 0, 0);
            heading.Controls.Add(new Label
            {
                Text = style == 18 ? "출력할 시험 대상을 선택하세요" : "시험 유형을 선택하세요",
                Dock = DockStyle.Fill,
                ForeColor = p.Foreground,
                Font = new Font("맑은 고딕", style == 18 ? 19 : 22, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            }, 1, 0);
            var headingSubtitle = new Label
            {
                Text = "대상을 선택하면 공정 선택과 출력 조건 확인 단계로 이동합니다.",
                Dock = DockStyle.Fill,
                ForeColor = p.Muted,
                Font = new Font("맑은 고딕", 9),
                TextAlign = ContentAlignment.MiddleLeft
            };
            heading.SetColumnSpan(headingSubtitle, 2);
            heading.Controls.Add(headingSubtitle, 0, 1);
            root.Controls.Add(heading, 0, 0);

            var cards = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = p.Background, Margin = new Padding(0, 8, 0, 8) };
            if (style == 4)
            {
                cards.ColumnCount = 2;
                cards.RowCount = 2;
                cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                cards.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
                cards.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            }
            else
            {
                cards.ColumnCount = 4;
                cards.RowCount = 1;
                for (int i = 0; i < 4; i++) cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
                cards.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            }
            root.Controls.Add(cards, 0, 1);

            Color general = style == 1 ? Color.FromArgb(0, 153, 132) : (style == 18 ? Color.FromArgb(34, 53, 60) : Color.White);
            Color electric = style == 1 ? Color.FromArgb(52, 104, 216) : (style == 18 ? Color.FromArgb(34, 53, 60) : Color.White);
            Color newcomer = style == 1 ? Color.FromArgb(91, 151, 47) : (style == 18 ? Color.FromArgb(34, 53, 60) : Color.White);
            Color global = style == 1 ? Color.FromArgb(201, 72, 107) : (style == 18 ? Color.FromArgb(34, 53, 60) : Color.White);
            Color text = style == 1 || style == 18 ? Color.White : p.Foreground;

            Control c1 = AudienceCard("general", "일반용", "GENERAL", "정기 작업자 시험지", general, text, Color.FromArgb(0, 153, 132), style);
            Control c2 = AudienceCard("electric", "전장용", "ELECTRICAL", "VDA 포함 시험지", electric, text, Color.FromArgb(52, 104, 216), style);
            Control c3 = AudienceCard("newcomer", "신입용", "NEWCOMER", "교육 연계 시험지", newcomer, text, Color.FromArgb(91, 151, 47), style);
            Control c4 = AudienceCard("global", "외국인용", "GLOBAL", "문제은행 추가 예정", global, text, Color.FromArgb(201, 72, 107), style);
            cards.Controls.Add(c1, 0, 0);
            cards.Controls.Add(c2, 1, 0);
            if (style == 4)
            {
                cards.Controls.Add(c3, 0, 1);
                cards.Controls.Add(c4, 1, 1);
            }
            else
            {
                cards.Controls.Add(c3, 2, 0);
                cards.Controls.Add(c4, 3, 0);
            }

            root.Controls.Add(new Label
            {
                Text = "✓  종이 시험 출력용 · 문제지와 답안지를 함께 출력합니다.",
                Dock = DockStyle.Fill,
                BackColor = p.Surface,
                ForeColor = p.Foreground,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(14, 0, 14, 0),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("맑은 고딕", 9, FontStyle.Bold)
            }, 0, 2);
            return page;
        }

        Control AudienceCard(string key, string title, string english, string subtitle, Color background, Color foreground, Color accent, int style)
        {
            var card = new Panel { Dock = DockStyle.Fill, BackColor = background, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(7), Cursor = Cursors.Hand };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, BackColor = background, Padding = new Padding(style == 18 ? 16 : 22) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, style == 18 ? 45 : 62));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            card.Controls.Add(layout);
            if (style == 18)
            {
                layout.Controls.Add(new Label { Text = "LINE  " + english, Dock = DockStyle.Fill, ForeColor = accent, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 8, FontStyle.Bold) }, 0, 0);
            }
            else
            {
                var iconHost = new Panel { Dock = DockStyle.Fill, BackColor = background };
                string iconKind = key == "general" ? "clipboard" : (key == "electric" ? "chip" : (key == "newcomer" ? "graduation" : "language"));
                Control icon = MakeLineIcon(iconKind, foreground, background, true);
                icon.Dock = DockStyle.None;
                icon.Left = 0;
                icon.Top = 5;
                icon.Width = 46;
                icon.Height = 46;
                iconHost.Controls.Add(icon);
                layout.Controls.Add(iconHost, 0, 0);
            }
            layout.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, ForeColor = foreground, TextAlign = style == 18 ? ContentAlignment.MiddleCenter : ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", style == 18 ? 18 : 21, FontStyle.Bold) }, 0, 1);
            layout.Controls.Add(new Label { Text = english + "\r\n" + subtitle, Dock = DockStyle.Fill, ForeColor = style == 1 ? Color.White : (style == 18 ? Color.FromArgb(190, 210, 214) : Color.FromArgb(91, 112, 113)), TextAlign = style == 18 ? ContentAlignment.MiddleCenter : ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 8) }, 0, 2);
            layout.Controls.Add(new Label { Text = "시험지 준비  →", Dock = DockStyle.Fill, ForeColor = style == 18 ? Color.White : foreground, TextAlign = ContentAlignment.BottomRight, Font = new Font("맑은 고딕", 10, FontStyle.Bold), Padding = new Padding(0, 0, 2, 2) }, 0, 3);
            EventHandler click = delegate { SelectAudience(key, title); };
            BindClick(card, click);
            return card;
        }

        Control MakeLineIcon(string kind, Color foreground, Color background, bool framed)
        {
            var icon = new Panel { Dock = DockStyle.Fill, BackColor = background, Margin = new Padding(0) };
            icon.Paint += delegate(object sender, PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                int size = Math.Max(18, Math.Min(icon.ClientSize.Width, icon.ClientSize.Height) - 4);
                var frame = new Rectangle((icon.ClientSize.Width - size) / 2, (icon.ClientSize.Height - size) / 2, size - 1, size - 1);
                using (var framePen = new Pen(Color.FromArgb(150, foreground), Math.Max(1f, size / 34f)))
                    if (framed) e.Graphics.DrawRectangle(framePen, frame);
                Rectangle glyph = framed ? Rectangle.Inflate(frame, -8, -8) : Rectangle.Inflate(frame, -4, -4);
                DrawLineIcon(e.Graphics, kind, glyph, foreground);
            };
            return icon;
        }

        static void DrawLineIcon(Graphics g, string kind, Rectangle r, Color color)
        {
            float stroke = Math.Max(1.4f, r.Width / 14f);
            using (var pen = new Pen(color, stroke) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                if (kind == "paper")
                {
                    var sheet = new Rectangle(r.Left + r.Width / 5, r.Top + 1, r.Width * 3 / 5, r.Height - 2);
                    g.DrawRectangle(pen, sheet);
                    int fold = Math.Max(4, r.Width / 5);
                    g.DrawLine(pen, sheet.Right - fold, sheet.Top, sheet.Right - fold, sheet.Top + fold);
                    g.DrawLine(pen, sheet.Right - fold, sheet.Top + fold, sheet.Right, sheet.Top + fold);
                    for (int i = 0; i < 3; i++)
                    {
                        int y = sheet.Top + fold + 4 + i * Math.Max(4, (sheet.Height - fold - 7) / 3);
                        g.DrawRectangle(pen, sheet.Left + 4, y, 2, 2);
                        g.DrawLine(pen, sheet.Left + 10, y + 1, sheet.Right - 4, y + 1);
                    }
                    return;
                }
                if (kind == "clipboard")
                {
                    var board = new Rectangle(r.Left + r.Width / 6, r.Top + r.Height / 7, r.Width * 2 / 3, r.Height * 5 / 7);
                    g.DrawRectangle(pen, board);
                    var clip = new Rectangle(r.Left + r.Width / 3, r.Top + 1, r.Width / 3, Math.Max(4, r.Height / 5));
                    g.DrawRectangle(pen, clip);
                    for (int i = 0; i < 3; i++)
                    {
                        int y = board.Top + 7 + i * Math.Max(5, (board.Height - 10) / 3);
                        g.DrawRectangle(pen, board.Left + 4, y, 2, 2);
                        g.DrawLine(pen, board.Left + 10, y + 1, board.Right - 4, y + 1);
                    }
                    return;
                }
                if (kind == "chip")
                {
                    var body = new Rectangle(r.Left + r.Width / 4, r.Top + r.Height / 4, r.Width / 2, r.Height / 2);
                    g.DrawRectangle(pen, body);
                    int[] offsets = { 1, 2, 3 };
                    foreach (int n in offsets)
                    {
                        int x = r.Left + r.Width * n / 4;
                        int y = r.Top + r.Height * n / 4;
                        g.DrawLine(pen, x, r.Top + 1, x, body.Top);
                        g.DrawLine(pen, x, body.Bottom, x, r.Bottom - 1);
                        g.DrawLine(pen, r.Left + 1, y, body.Left, y);
                        g.DrawLine(pen, body.Right, y, r.Right - 1, y);
                    }
                    g.DrawEllipse(pen, body.Left + body.Width / 3, body.Top + body.Height / 3, Math.Max(3, body.Width / 3), Math.Max(3, body.Height / 3));
                    return;
                }
                if (kind == "graduation")
                {
                    var cap = new[]
                    {
                        new Point(r.Left + r.Width / 2, r.Top + 2),
                        new Point(r.Right - 2, r.Top + r.Height / 3),
                        new Point(r.Left + r.Width / 2, r.Top + r.Height * 2 / 3),
                        new Point(r.Left + 2, r.Top + r.Height / 3)
                    };
                    g.DrawPolygon(pen, cap);
                    int baseY = r.Top + r.Height * 2 / 3;
                    g.DrawArc(pen, r.Left + r.Width / 4, baseY - r.Height / 8, r.Width / 2, r.Height / 4, 0, 180);
                    g.DrawLine(pen, r.Right - 3, r.Top + r.Height / 3, r.Right - 3, r.Bottom - 4);
                    g.DrawEllipse(pen, r.Right - 5, r.Bottom - 5, 4, 4);
                    return;
                }

                var globe = new Rectangle(r.Left + 1, r.Top + r.Height / 5, r.Width * 3 / 5, r.Height * 3 / 5);
                g.DrawEllipse(pen, globe);
                g.DrawEllipse(pen, globe.Left + globe.Width / 3, globe.Top, globe.Width / 3, globe.Height);
                g.DrawLine(pen, globe.Left + 2, globe.Top + globe.Height / 2, globe.Right - 2, globe.Top + globe.Height / 2);
                using (var font = new Font("Segoe UI", Math.Max(7f, r.Height / 3f), FontStyle.Bold))
                using (var brush = new SolidBrush(color))
                    g.DrawString("A", font, brush, r.Right - r.Width / 3, r.Bottom - r.Height / 2 - 2);
            }
        }

        void BindClick(Control control, EventHandler handler)
        {
            control.Click += handler;
            control.Cursor = Cursors.Hand;
            foreach (Control child in control.Controls) BindClick(child, handler);
        }

        Panel BuildThreeColumnWorkspace(UiPalette p, bool allowBack)
        {
            var page = new Panel { Dock = DockStyle.Fill, BackColor = p.Background };
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = p.Background, Padding = new Padding(0, 8, 0, 0) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            page.Controls.Add(root);

            var top = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, BackColor = p.Surface, Padding = new Padding(6) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, allowBack ? 112 : 4));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
            if (allowBack) top.Controls.Add(MakeButton("← 대상 선택", delegate { ShowLanding(); }, p.SurfaceAlt, p.Foreground, p.Border, 9), 0, 0);
            audienceLabel = new Label { Text = "시험지 출력 준비", Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12, 0, 0, 0), Font = new Font("맑은 고딕", 14, FontStyle.Bold) };
            top.Controls.Add(audienceLabel, 1, 0);
            top.Controls.Add(new Label { Text = "종이 시험 출력용", Dock = DockStyle.Fill, ForeColor = p.Accent, TextAlign = ContentAlignment.MiddleRight, Font = new Font("맑은 고딕", 9, FontStyle.Bold), Padding = new Padding(0, 0, 10, 0) }, 2, 0);
            root.Controls.Add(top, 0, 0);

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, BackColor = p.Background, Padding = new Padding(0, 8, 0, 0) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            root.Controls.Add(body, 0, 1);
            body.Controls.Add(BuildBankSection(p, false), 0, 0);
            body.Controls.Add(BuildConditionSection(p, true, false), 1, 0);
            body.Controls.Add(BuildPreviewSection(p), 2, 0);
            return page;
        }

        Panel BuildCarbonWorkspace(UiPalette p)
        {
            var page = new Panel { Dock = DockStyle.Fill, BackColor = p.Background, Padding = new Padding(18) };
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = p.Background };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
            page.Controls.Add(body);

            var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = p.Background, Margin = new Padding(0, 0, 16, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 92));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var heading = new Panel { Dock = DockStyle.Fill, BackColor = p.Background };
            heading.Controls.Add(new Label { Text = "EXAM PRINT", Left = 0, Top = 4, Width = 220, Height = 22, ForeColor = p.Accent, Font = new Font("Segoe UI", 8, FontStyle.Bold) });
            heading.Controls.Add(new Label { Text = "공정 시험지 출력", Left = 0, Top = 26, Width = 420, Height = 38, ForeColor = p.Foreground, Font = new Font("맑은 고딕", 20, FontStyle.Bold) });
            heading.Controls.Add(new Label { Text = "공정을 선택하고 종이 시험지를 생성한 뒤 출력합니다.", Left = 0, Top = 67, Width = 540, Height = 22, ForeColor = p.Muted, Font = new Font("맑은 고딕", 8) });
            left.Controls.Add(heading, 0, 0);
            left.Controls.Add(BuildBankSection(p, true), 0, 1);
            body.Controls.Add(left, 0, 0);
            body.Controls.Add(BuildConditionSection(p, false, true), 1, 0);

            var hidden = new Panel { Visible = false, Width = 1, Height = 1 };
            preview = new Preview { Dock = DockStyle.Fill };
            pageLabel = new Label { Text = "1 / 1" };
            hidden.Controls.Add(preview);
            hidden.Controls.Add(pageLabel);
            page.Controls.Add(hidden);
            return page;
        }

        Control BuildBankSection(UiPalette p, bool carbon)
        {
            var box = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = p.Surface, Margin = new Padding(5), Padding = new Padding(12) };
            box.RowStyles.Add(new RowStyle(SizeType.Absolute, carbon ? 42 : 48));
            box.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            box.Controls.Add(new Label
            {
                Text = carbon ? "공정 목록" : "1  공정 선택",
                Dock = DockStyle.Fill,
                ForeColor = p.Foreground,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("맑은 고딕", carbon ? 11 : 13, FontStyle.Bold)
            }, 0, 0);
            grid = CreateBankGrid(p, carbon);
            box.Controls.Add(grid, 0, 1);
            return box;
        }

        DataGridView CreateBankGrid(UiPalette p, bool carbon)
        {
            var value = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeColumns = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                BackgroundColor = p.Surface,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                GridColor = p.Border,
                Tag = carbon ? "carbon" : "counts"
            };
            value.ColumnHeadersDefaultCellStyle.BackColor = carbon ? Color.FromArgb(54, 54, 54) : p.SurfaceAlt;
            value.ColumnHeadersDefaultCellStyle.ForeColor = carbon ? Color.White : p.Foreground;
            value.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 8, FontStyle.Bold);
            value.DefaultCellStyle.BackColor = p.Surface;
            value.DefaultCellStyle.ForeColor = p.Foreground;
            value.DefaultCellStyle.SelectionBackColor = carbon ? Color.FromArgb(221, 232, 249) : Color.FromArgb(216, 235, 230);
            value.DefaultCellStyle.SelectionForeColor = p.Foreground;
            value.DefaultCellStyle.Font = new Font("맑은 고딕", carbon ? 8 : 9);
            value.RowTemplate.Height = carbon ? 42 : 30;
            value.ColumnHeadersHeight = carbon ? 38 : 30;
            if (carbon)
            {
                value.Columns.Add("no", "NO");
                value.Columns.Add("name", "공정명");
                value.Columns.Add("audience", "대상");
                value.Columns.Add("state", "상태");
                value.Columns[0].FillWeight = 30;
                value.Columns[1].FillWeight = 220;
                value.Columns[2].FillWeight = 70;
                value.Columns[3].FillWeight = 68;
            }
            else
            {
                value.Columns.Add("name", "공정명");
                value.Columns.Add("vda", "VDA");
                value.Columns.Add("choice", "객관식");
                value.Columns.Add("subjective", "주관식");
                value.Columns[0].FillWeight = 240;
                value.Columns[1].FillWeight = 52;
                value.Columns[2].FillWeight = 70;
                value.Columns[3].FillWeight = 70;
            }
            value.CellFormatting += delegate(object s, DataGridViewCellFormattingEventArgs e)
            {
                if (e.RowIndex >= 0 && value.Columns[e.ColumnIndex].Name == "name" && e.Value != null)
                    value.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = e.Value.ToString();
            };
            value.SelectionChanged += delegate { RefreshSelected(); };
            return value;
        }

        Control BuildConditionSection(UiPalette p, bool inlinePreview, bool carbon)
        {
            var outer = new Panel { Dock = DockStyle.Fill, BackColor = p.Surface, Margin = new Padding(5), Padding = new Padding(0) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, BackColor = p.Surface, Padding = new Padding(14) };
            outer.Controls.Add(flow);

            var titleRow = new TableLayoutPanel { Width = 360, Height = 38, ColumnCount = 1, Margin = new Padding(0, 0, 0, 3) };
            titleRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            titleRow.Controls.Add(new Label { Text = carbon ? "SELECTED" : "2  출력 조건", Dock = DockStyle.Fill, ForeColor = carbon ? p.Accent : p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Font = new Font(carbon ? "Segoe UI" : "맑은 고딕", carbon ? 8 : 13, FontStyle.Bold) }, 0, 0);
            flow.Controls.Add(titleRow);

            selectedBankNameLabel = new Label { Text = "공정을 선택하세요", Width = 360, Height = 34, ForeColor = p.Foreground, Font = new Font("맑은 고딕", carbon ? 16 : 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, Margin = new Padding(0, 0, 0, 0) };
            flow.Controls.Add(selectedBankNameLabel);
            selectedBankMetaLabel = new Label { Text = "", Width = 360, Height = 24, ForeColor = p.Muted, Font = new Font("맑은 고딕", 8), TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, Margin = new Padding(0, 0, 0, 3) };
            flow.Controls.Add(selectedBankMetaLabel);

            var header = new TableLayoutPanel { Width = 360, Height = 22, ColumnCount = 3, Margin = new Padding(0) };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            header.Controls.Add(new Label { Text = "항목", Dock = DockStyle.Fill, ForeColor = p.Muted, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 8, FontStyle.Bold) }, 0, 0);
            header.Controls.Add(new Label { Text = "문제수", Dock = DockStyle.Fill, ForeColor = p.Muted, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("맑은 고딕", 8, FontStyle.Bold) }, 1, 0);
            header.Controls.Add(new Label { Text = "점수", Dock = DockStyle.Fill, ForeColor = p.Muted, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("맑은 고딕", 8, FontStyle.Bold) }, 2, 0);
            flow.Controls.Add(header);

            vdaRow = ReadoutRow("VDA", out vdaCount, out vdaScore, 2, 2.5m, p);
            flow.Controls.Add(vdaRow);
            flow.Controls.Add(ReadoutRow("객관식", out choiceCount, out choiceScore, 20, 4m, p));
            flow.Controls.Add(ReadoutRow("주관식", out subjectiveCount, out subjectiveScore, 4, 5m, p));
            targetScore = ReadoutNum(0, 1000, 100, 0, p);
            flow.Controls.Add(ReadoutTargetRow(p));

            totalLabel = new Label { Text = "TOTAL 100", Width = 360, Height = 42, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(0, 142, 99), BackColor = p.SurfaceAlt, Font = new Font("Segoe UI", 19, FontStyle.Bold), Margin = new Padding(0, 4, 0, 3) };
            flow.Controls.Add(totalLabel);
            flow.Controls.Add(ReadoutInputRow("성명", out userNameBox, "", p));
            flow.Controls.Add(ReadoutInputRow("평가 일시", out evalDateBox, DateTime.Today.ToString("yyyy.MM.dd"), p));
            answerCheck = new CheckBox { Text = "화면 미리보기에 답안 표시", Width = 360, Height = 24, ForeColor = p.Foreground, Font = new Font("맑은 고딕", 8), Margin = new Padding(0, 3, 0, 0) };
            answerCheck.CheckedChanged += delegate
            {
                if (doc != null)
                {
                    doc.ShowAnswers = answerCheck.Checked;
                    if (preview != null) preview.Invalidate();
                }
            };
            flow.Controls.Add(answerCheck);

            generationLabel = new Label { Text = "공정과 출력 정보를 확인하세요.", Width = 360, Height = 27, ForeColor = p.Muted, Font = new Font("맑은 고딕", 8), TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, Margin = new Padding(0, 1, 0, 2) };
            flow.Controls.Add(generationLabel);
            EventHandler createHandler = carbon ? new EventHandler(GenerateAndShowScreenPreview) : new EventHandler(Generate);
            flow.Controls.Add(MakeActionButton("시험지 만들기", createHandler, p.Accent, p.AccentText));
            flow.Controls.Add(MakeActionButton("출력 미리보기", delegate { PrintPreviewProblemAndAnswer(); }, p.SurfaceAlt, p.Foreground));
            flow.Controls.Add(MakeActionButton("문제 + 답안 프린트", delegate { PrintProblemAndAnswer(); }, carbon ? p.Accent : p.Header, Color.White));

            flow.Resize += delegate
            {
                int width = Math.Max(260, flow.ClientSize.Width - 34);
                foreach (Control child in flow.Controls) child.Width = width;
            };
            return outer;
        }

        Control BuildPreviewSection(UiPalette p)
        {
            var box = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = p.SurfaceAlt, Margin = new Padding(5), Padding = new Padding(12) };
            box.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            box.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            box.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4 };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44));
            top.Controls.Add(new Label { Text = "3  시험지 화면 확인", Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 13, FontStyle.Bold) }, 0, 0);
            top.Controls.Add(MakeButton("−", delegate { preview.ZoomBy(1f / 1.12f); }, p.Surface, p.Foreground, p.Border, 11), 1, 0);
            top.Controls.Add(MakeButton("맞춤", delegate { preview.Fit(); }, p.Surface, p.Foreground, p.Border, 8), 2, 0);
            top.Controls.Add(MakeButton("+", delegate { preview.ZoomBy(1.12f); }, p.Surface, p.Foreground, p.Border, 11), 3, 0);
            box.Controls.Add(top, 0, 0);
            preview = new Preview { Dock = DockStyle.Fill, BackColor = Color.FromArgb(12, 24, 34) };
            box.Controls.Add(preview, 0, 1);
            var nav = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(0, 4, 0, 0) };
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            nav.Controls.Add(MakeButton("이전", delegate { preview.MovePage(-1); UpdatePage(); }, p.Surface, p.Foreground, p.Border, 8), 0, 0);
            pageLabel = new Label { Text = "1 / 1", Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            nav.Controls.Add(pageLabel, 1, 0);
            nav.Controls.Add(MakeButton("다음", delegate { preview.MovePage(1); UpdatePage(); }, p.Surface, p.Foreground, p.Border, 8), 2, 0);
            box.Controls.Add(nav, 0, 2);
            return box;
        }

        Control ReadoutRow(string name, out NumericUpDown count, out NumericUpDown score, int countValue, decimal scoreValue, UiPalette p)
        {
            var row = new TableLayoutPanel { Width = 360, Height = 34, ColumnCount = 3, Margin = new Padding(0, 1, 0, 1) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            row.Controls.Add(new Label { Text = name, Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 10, FontStyle.Bold) }, 0, 0);
            count = ReadoutNum(0, 500, countValue, 0, p);
            score = ReadoutNum(0, 100, scoreValue, 1, p);
            row.Controls.Add(count, 1, 0);
            row.Controls.Add(score, 2, 0);
            return row;
        }

        Control ReadoutTargetRow(UiPalette p)
        {
            var row = new TableLayoutPanel { Width = 360, Height = 34, ColumnCount = 2, Margin = new Padding(0, 1, 0, 1) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 71));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            row.Controls.Add(new Label { Text = "목표 점수", Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 10, FontStyle.Bold) }, 0, 0);
            row.Controls.Add(targetScore, 1, 0);
            return row;
        }

        Control ReadoutInputRow(string name, out TextBox box, string value, UiPalette p)
        {
            var row = new TableLayoutPanel { Width = 360, Height = 36, ColumnCount = 2, Margin = new Padding(0, 1, 0, 1) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            row.Controls.Add(new Label { Text = name, Dock = DockStyle.Fill, ForeColor = p.Foreground, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("맑은 고딕", 9, FontStyle.Bold) }, 0, 0);
            box = new TextBox { Text = value, Dock = DockStyle.Fill, BackColor = p.Input, ForeColor = p.Foreground, BorderStyle = BorderStyle.FixedSingle, Font = new Font("맑은 고딕", 10, FontStyle.Bold), TextAlign = HorizontalAlignment.Center, Margin = new Padding(0, 6, 0, 5) };
            box.TextChanged += delegate
            {
                if (doc != null)
                {
                    doc.UserName = userNameBox.Text.Trim();
                    doc.EvalDate = evalDateBox.Text.Trim();
                    if (preview != null) preview.Invalidate();
                }
            };
            row.Controls.Add(box, 1, 0);
            return row;
        }

        NumericUpDown ReadoutNum(decimal min, decimal max, decimal value, int decimals, UiPalette p)
        {
            var number = Num(min, max, value, decimals);
            number.BackColor = p.Input;
            number.ForeColor = p.Foreground;
            number.BorderStyle = BorderStyle.FixedSingle;
            number.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            number.Margin = new Padding(4, 5, 0, 4);
            return number;
        }

        Button MakeButton(string text, EventHandler handler, Color background, Color foreground, Color border, float fontSize)
        {
            var value = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                BackColor = background,
                ForeColor = foreground,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("맑은 고딕", fontSize, FontStyle.Bold),
                Margin = new Padding(3),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            value.FlatAppearance.BorderColor = border;
            value.FlatAppearance.BorderSize = 1;
            value.Click += handler;
            return value;
        }

        Control MakeHeaderAction(string text, EventHandler handler, UiPalette p)
        {
            var value = new Label
            {
                Text = "▣  " + text,
                Dock = DockStyle.Fill,
                BackColor = p.Surface,
                ForeColor = p.Foreground,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", 9, FontStyle.Bold),
                Margin = new Padding(3),
                Cursor = Cursors.Hand
            };
            value.Click += handler;
            return value;
        }

        Button MakeActionButton(string text, EventHandler handler, Color background, Color foreground)
        {
            var value = MakeButton(text, handler, background, foreground, background, 9);
            value.Width = 360;
            value.Height = 36;
            value.Margin = new Padding(0, 3, 0, 1);
            return value;
        }

        void SelectAudience(string key, string title)
        {
            audienceFilter = key;
            PopulateBankGrid();
            if (visibleBanks.Count == 0)
            {
                MessageBox.Show(this, title + " 문제은행은 아직 등록되지 않았습니다.\r\n엑셀에 해당 공정이 추가되면 자동으로 표시됩니다.", "문제은행 확인", MessageBoxButtons.OK, MessageBoxIcon.Information);
                audienceFilter = "";
                PopulateBankGrid();
                return;
            }
            if (audienceLabel != null) audienceLabel.Text = title + " 시험지 출력 준비";
            if (stepLabel != null) stepLabel.Text = "3 / 4 단계";
            if (landingPage != null) landingPage.Visible = false;
            if (workspacePage != null)
            {
                workspacePage.Visible = true;
                workspacePage.BringToFront();
            }
        }

        void ShowLanding()
        {
            if (landingPage == null) return;
            audienceFilter = "";
            PopulateBankGrid();
            if (stepLabel != null) stepLabel.Text = "2 / 4 단계";
            if (workspacePage != null) workspacePage.Visible = false;
            landingPage.Visible = true;
            landingPage.BringToFront();
        }

        void PopulateBankGrid()
        {
            if (grid == null) return;
            visibleBanks.Clear();
            foreach (Bank bank in banks)
                if (BankMatchesAudience(bank, audienceFilter)) visibleBanks.Add(bank);

            grid.Rows.Clear();
            bool carbon = Convert.ToString(grid.Tag) == "carbon";
            for (int i = 0; i < visibleBanks.Count; i++)
            {
                Bank bank = visibleBanks[i];
                if (carbon)
                    grid.Rows.Add((i + 1).ToString("00"), bank.Name, BankAudience(bank), "출력 가능");
                else
                    grid.Rows.Add(bank.Name, bank.Vda, bank.Choice, bank.Subjective);
            }
            if (grid.Rows.Count > 0)
            {
                grid.ClearSelection();
                grid.Rows[0].Selected = true;
                grid.CurrentCell = grid.Rows[0].Cells[0];
            }
            RefreshSelected();
        }

        static bool BankMatchesAudience(Bank bank, string filter)
        {
            if (bank == null || string.IsNullOrEmpty(filter)) return bank != null;
            string name = (bank.Name ?? "").Replace(" ", "");
            if (filter == "electric") return name.IndexOf("전장용", StringComparison.Ordinal) >= 0;
            if (filter == "newcomer") return name.IndexOf("신입용", StringComparison.Ordinal) >= 0;
            if (filter == "global") return name.IndexOf("외국인", StringComparison.Ordinal) >= 0 || name.IndexOf("외국", StringComparison.Ordinal) >= 0;
            return name.IndexOf("일반용", StringComparison.Ordinal) >= 0;
        }

        static string BankAudience(Bank bank)
        {
            string name = bank == null ? "" : bank.Name ?? "";
            if (name.IndexOf("전장용", StringComparison.Ordinal) >= 0) return "전장용";
            if (name.IndexOf("신입용", StringComparison.Ordinal) >= 0) return "신입용";
            if (name.IndexOf("외국", StringComparison.Ordinal) >= 0) return "외국인용";
            return "일반용";
        }

        void GenerateAndShowScreenPreview(object sender, EventArgs e)
        {
            Generate(sender, e);
            if (doc != null) ShowScreenPreview();
        }

        void ShowScreenPreview()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "시험지 화면 확인 · 실제 출력은 [출력 미리보기]에서 확인";
                dialog.Width = 1180;
                dialog.Height = 840;
                dialog.MinimumSize = new Size(900, 650);
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.BackColor = Color.FromArgb(20, 32, 38);
                var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(12), BackColor = dialog.BackColor };
                root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
                dialog.Controls.Add(root);
                var screen = new Preview { Dock = DockStyle.Fill, Doc = doc, BackColor = Color.FromArgb(9, 20, 30) };
                root.Controls.Add(screen, 0, 0);
                var nav = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5 };
                nav.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
                nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                nav.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
                nav.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
                nav.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
                var label = new Label { Text = "1 / " + screen.PageCount, Dock = DockStyle.Fill, ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
                nav.Controls.Add(MakeButton("이전", delegate { screen.MovePage(-1); label.Text = (screen.PageIndex + 1) + " / " + screen.PageCount; }, Color.FromArgb(45, 62, 69), Color.White, Color.FromArgb(78, 96, 102), 9), 0, 0);
                nav.Controls.Add(label, 1, 0);
                nav.Controls.Add(MakeButton("맞춤", delegate { screen.Fit(); }, Color.FromArgb(45, 62, 69), Color.White, Color.FromArgb(78, 96, 102), 9), 2, 0);
                nav.Controls.Add(MakeButton("다음", delegate { screen.MovePage(1); label.Text = (screen.PageIndex + 1) + " / " + screen.PageCount; }, Color.FromArgb(45, 62, 69), Color.White, Color.FromArgb(78, 96, 102), 9), 3, 0);
                nav.Controls.Add(MakeButton("닫기", delegate { dialog.Close(); }, Color.White, Color.FromArgb(25, 37, 42), Color.White, 9), 4, 0);
                root.Controls.Add(nav, 0, 1);
                dialog.ShowDialog(this);
            }
        }

        internal void PrepareWorkspaceForTest()
        {
            if (UiVariant != 31) SelectAudience("general", "일반용");
        }

        internal void RunWorkflowSelfTest(string workbook)
        {
            if (!string.IsNullOrWhiteSpace(workbook))
            {
                pathBox.Text = workbook;
                LoadBanks();
            }
            foreach (string key in new[] { "general", "electric", "newcomer" })
            {
                audienceFilter = key;
                PopulateBankGrid();
                Bank bank = SelectedBank();
                if (bank == null) throw new InvalidOperationException("Missing workflow bank: " + key);
                decimal total = vdaCount.Value * vdaScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value;
                if (total != targetScore.Value) throw new InvalidOperationException("Workflow score mismatch: " + key);
                var testDoc = new ExamDoc { Bank = bank, UserName = "TEST", EvalDate = DateTime.Today.ToString("yyyy.MM.dd") };
                Pick(testDoc, bank, "VDA", (int)vdaCount.Value, (double)vdaScore.Value);
                Pick(testDoc, bank, "객관식", (int)choiceCount.Value, (double)choiceScore.Value);
                Pick(testDoc, bank, "주관식", (int)subjectiveCount.Value, (double)subjectiveScore.Value);
                int expected = (int)(vdaCount.Value + choiceCount.Value + subjectiveCount.Value);
                if (testDoc.Questions.Count != expected) throw new InvalidOperationException("Workflow question count mismatch: " + key);
                double score = testDoc.Questions.Sum(q => q.Score);
                if (Math.Abs(score - (double)targetScore.Value) > 0.001) throw new InvalidOperationException("Workflow selected score mismatch: " + key);
            }
            audienceFilter = "";
            PopulateBankGrid();
        }
        void Browse(object s, EventArgs e) { using (var d = new OpenFileDialog { Filter = "Excel files (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|All files (*.*)|*.*", Title = "\uBB38\uC81C\uC740\uD589 \uC5D1\uC140 \uC120\uD0DD" }) { if (File.Exists(pathBox.Text)) d.InitialDirectory = Path.GetDirectoryName(pathBox.Text); if (d.ShowDialog(this) == DialogResult.OK) { pathBox.Text = d.FileName; SaveLastWorkbook(d.FileName); LoadBanks(); } } }
        void LoadBanks()
        {
            try
            {
                Status("LOADING", false); doc = null; if (preview != null) { preview.Doc = null; preview.PageIndex = 0; preview.Invalidate(); } banks.Clear(); banks.AddRange(reader.Load(pathBox.Text)); SaveLastWorkbook(pathBox.Text);
                PopulateBankGrid();
                RefreshSelected(); UpdateTotal(); UpdatePage(); Status("READY", false);
            }
            catch (Exception ex) { Status("ERROR", true); MessageBox.Show(this, ex.Message, "\uC624\uB958", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        Bank SelectedBank() { return grid.CurrentRow == null || grid.CurrentRow.Index < 0 || grid.CurrentRow.Index >= visibleBanks.Count ? null : visibleBanks[grid.CurrentRow.Index]; }
        void OpenExamSettings(object sender, EventArgs e)
        {
            using (var password = new SettingsPasswordForm())
                if (password.ShowDialog(this) != DialogResult.OK) return;

            using (var dialog = new ExamSettingsForm(examSettings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    ExamSettingsStore.Save(dialog.Value);
                    examSettings = dialog.Value.Clone();
                    ApplyModeSettings(SelectedBank());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "\uC124\uC815 \uC800\uC7A5 \uC624\uB958", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        void ApplyModeSettings(Bank bank)
        {
            if (vdaRow == null || vdaCount == null || targetScore == null) return;
            bool electric = IsElectricBank(bank);
            ExamModeSettings mode = electric ? examSettings.Electric : examSettings.General;
            vdaRow.Visible = true;
            foreach (Control child in vdaRow.Controls)
                child.Visible = electric;
            SetValue(vdaCount, electric ? mode.VdaCount : 0);
            SetValue(vdaScore, mode.VdaScore);
            SetValue(choiceCount, mode.ChoiceCount);
            SetValue(choiceScore, mode.ChoiceScore);
            SetValue(subjectiveCount, mode.SubjectiveCount);
            SetValue(subjectiveScore, mode.SubjectiveScore);
            SetValue(targetScore, mode.TargetScore);
            UpdateTotal();
        }
        static bool IsElectricBank(Bank bank)
        {
            return bank != null && bank.Name.IndexOf("\uC804\uC7A5\uC6A9", StringComparison.Ordinal) >= 0;
        }
        static void SetValue(NumericUpDown number, decimal value)
        {
            number.Value = Math.Max(number.Minimum, Math.Min(number.Maximum, value));
        }
        void RefreshSelected()
        {
            Bank bank = SelectedBank();
            ApplyModeSettings(bank);
            if (selectedBankNameLabel != null) selectedBankNameLabel.Text = bank == null ? "공정을 선택하세요" : bank.Name;
            if (selectedBankMetaLabel != null)
                selectedBankMetaLabel.Text = bank == null ? "" : string.Format("{0} · VDA {1} / 객관식 {2} / 주관식 {3}", BankAudience(bank), bank.Vda, bank.Choice, bank.Subjective);
            if (generationLabel != null) generationLabel.Text = bank == null ? "공정을 먼저 선택하세요." : "출력 조건과 성명을 확인하세요.";
        }
        void UpdateTotal() { if (totalLabel == null) return; decimal total = vdaCount.Value * vdaScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value; totalLabel.Text = "TOTAL " + total.ToString("0.##"); totalLabel.ForeColor = total == targetScore.Value ? Color.FromArgb(32, 208, 132) : Color.FromArgb(246, 180, 75); }
        void Generate(object s, EventArgs e)
        {
            try
            {
                Bank b = SelectedBank(); if (b == null) throw new InvalidOperationException("\uACF5\uC815\uC744 \uC120\uD0DD\uD558\uC138\uC694.");
                decimal total = vdaCount.Value * vdaScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value;
                if (total != targetScore.Value) throw new InvalidOperationException(string.Format("\uCD1D\uC810\uC774 \uBAA9\uD45C \uC810\uC218\uC640 \uB2E4\uB985\uB2C8\uB2E4. \uD604\uC7AC {0:0.##}\uC810 / \uBAA9\uD45C {1:0.##}\uC810", total, targetScore.Value));
                doc = new ExamDoc { Bank = b, ShowAnswers = answerCheck.Checked, UserName = userNameBox.Text.Trim(), EvalDate = evalDateBox.Text.Trim() };
                Pick(doc, b, "VDA", (int)vdaCount.Value, (double)vdaScore.Value);
                Pick(doc, b, "\uAC1D\uAD00\uC2DD", (int)choiceCount.Value, (double)choiceScore.Value);
                Pick(doc, b, "\uC8FC\uAD00\uC2DD", (int)subjectiveCount.Value, (double)subjectiveScore.Value);
                preview.Doc = doc; preview.PageIndex = 0; preview.Invalidate(); UpdatePage(); Status("READY", false);
                if (generationLabel != null) generationLabel.Text = string.Format("{0} · {1}문제 · {2:0.##}점 생성 완료", b.Name, doc.Questions.Count, total);
            }
            catch (Exception ex) { Status("ERROR", true); MessageBox.Show(this, ex.Message, "\uC624\uB958", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        void Pick(ExamDoc d, Bank b, string type, int count, double score)
        {
            var cand = b.Questions.Where(q => q.Type == type).OrderBy(q => random.Next()).Take(count).ToList();
            if (cand.Count < count)
            {
                string legacy = type == "VDA" && b.LegacyCommon > 0
                    ? string.Format("\n\uC5D1\uC140\uC5D0 '\uACF5\uD1B5' {0}\uBB38\uD56D\uC774 \uB0A8\uC544 \uC788\uC2B5\uB2C8\uB2E4. VDA\uB85C \uBCC0\uACBD\uD558\uC138\uC694.", b.LegacyCommon)
                    : "";
                throw new InvalidOperationException(string.Format("{0} \uBB38\uC81C\uAC00 \uBD80\uC871\uD569\uB2C8\uB2E4. \uC694\uCCAD {1}\uBB38\uD56D / \uBCF4\uC720 {2}\uBB38\uD56D{3}", type, count, cand.Count, legacy));
            }
            foreach (var q in cand) { var n = new Question { SheetName = q.SheetName, No = q.No, Type = q.Type, Text = q.Text, Answer = q.Answer, Score = score }; n.Images.AddRange(q.Images); d.Questions.Add(n); }
        }
        void UpdatePage() { if (pageLabel != null) pageLabel.Text = (preview.PageIndex + 1) + " / " + preview.PageCount; }
        void EnsureDoc() { if (doc == null) Generate(null, EventArgs.Empty); if (doc == null) throw new InvalidOperationException("\uBBF8\uB9AC\uBCF4\uAE30\uB97C \uBA3C\uC800 \uC0DD\uC131\uD558\uC138\uC694."); doc.UserName = userNameBox.Text.Trim(); doc.EvalDate = evalDateBox.Text.Trim(); }
        void EnsurePrintInfo()
        {
            if (userNameBox.Text.Trim().Length == 0)
                throw new InvalidOperationException("\uC131\uBA85\uC744 \uC785\uB825\uD574\uC57C \uCD9C\uB825\uD560 \uC218 \uC788\uC2B5\uB2C8\uB2E4.");
            if (evalDateBox.Text.Trim().Length == 0)
                throw new InvalidOperationException("\uD3C9\uAC00 \uC77C\uC2DC\uB97C \uC785\uB825\uD574\uC57C \uCD9C\uB825\uD560 \uC218 \uC788\uC2B5\uB2C8\uB2E4.");
        }
        void PrintProblem(bool answers) { try { EnsureDoc(); PrintPages(doc.Bank.Name, preview.QuestionPageCount, (g, r, p) => preview.DrawPrint(g, r, p, answers)); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "\uC624\uB958"); } }
        void PrintPreviewProblemAndAnswer() { try { EnsureDoc(); EnsurePrintInfo(); int pc = preview.QuestionPageCount; PreviewPages("OJT \uC2DC\uD5D8\uC9C0 + \uB2F5\uC548\uC9C0", pc + 1, (g, r, p) => { if (p < pc) preview.DrawPrint(g, r, p, false); else preview.DrawAnswerPrint(g, r); }); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "\uC624\uB958"); } }
        void PrintProblemAndAnswer() { try { EnsureDoc(); EnsurePrintInfo(); int pc = preview.QuestionPageCount; PrintPages("OJT \uC2DC\uD5D8\uC9C0 + \uB2F5\uC548\uC9C0", pc + 1, (g, r, p) => { if (p < pc) preview.DrawPrint(g, r, p, false); else preview.DrawAnswerPrint(g, r); }); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "\uC624\uB958"); } }
        void PrintAnswerOnly() { try { EnsureDoc(); PrintPages("OJT \uB2F5\uC548\uC9C0", 1, (g, r, p) => DrawAnswer(g, r)); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "\uC624\uB958"); } }
        delegate void DrawPage(Graphics g, Rectangle r, int p);
        void PrintPages(string name, int pages, DrawPage draw)
        {
            using (var pd = new PrintDocument())
            {
                pd.DocumentName = name; pd.DefaultPageSettings.Landscape = false; pd.DefaultPageSettings.Margins = new Margins(18, 18, 18, 18);
                int page = 0; pd.PrintPage += delegate(object s, PrintPageEventArgs e) { draw(e.Graphics, e.MarginBounds, page); page++; e.HasMorePages = page < pages; };
                pd.BeginPrint += delegate { page = 0; };
                using (var dlg = new PrintDialog { Document = pd, UseEXDialog = true }) if (dlg.ShowDialog(this) == DialogResult.OK) pd.Print();
            }
        }
        void PreviewPages(string name, int pages, DrawPage draw)
        {
            using (var pd = new PrintDocument())
            {
                pd.DocumentName = name; pd.DefaultPageSettings.Landscape = false; pd.DefaultPageSettings.Margins = new Margins(18, 18, 18, 18);
                int page = 0; pd.PrintPage += delegate(object s, PrintPageEventArgs e) { draw(e.Graphics, e.MarginBounds, page); page++; e.HasMorePages = page < pages; };
                pd.BeginPrint += delegate { page = 0; };
                using (var dlg = new PrintPreviewDialog { Document = pd, Width = 1200, Height = 850, StartPosition = FormStartPosition.CenterParent, UseAntiAlias = true })
                {
                    PrintPreviewControl previewControl = dlg.PrintPreviewControl;
                    using (var wheelFilter = new PreviewWheelMessageFilter(dlg, previewControl, pages))
                        dlg.ShowDialog(this);
                }
            }
        }
        void DrawAnswer(Graphics g, Rectangle bounds)
        {
            DrawAnswerTable(g, bounds, doc.Bank.Name, doc.UserName, doc.EvalDate, doc.Questions, false);
        }
        internal static void DrawAnswerTable(Graphics g, Rectangle bounds, string titleText, string userName, string evalDate, List<Question> questions, bool scaled)
        {
            g.FillRectangle(Brushes.White, bounds);
            int titleSize = scaled ? Math.Max(14, bounds.Width / 42) : 22;
            int headSize = scaled ? Math.Max(9, bounds.Width / 70) : 13;
            int fontSize = scaled ? Math.Max(8, bounds.Width / 82) : 12;
            using (var title = new Font("\uB9D1\uC740 \uACE0\uB515", titleSize, FontStyle.Regular))
            using (var head = new Font("\uB9D1\uC740 \uACE0\uB515", headSize, FontStyle.Regular))
            using (var font = new Font("\uB9D1\uC740 \uACE0\uB515", fontSize, FontStyle.Regular))
            using (var gridPen = new Pen(Color.Black, 1))
            {
                int padX = Math.Max(4, bounds.Width / 120);
                int padY = Math.Max(4, bounds.Height / 80);
                int x = bounds.Left + padX;
                int y = bounds.Top + padY;
                int usableW = bounds.Width - padX * 2;
                g.DrawString(titleText ?? "", title, Brushes.Blue, x, y);
                y += Math.Max(scaled ? 30 : 48, title.Height + (scaled ? 8 : 14));

                int infoH = Math.Max(scaled ? 44 : 58, font.Height * 2 + 12);
                int labelW = Math.Max(scaled ? 54 : 82, usableW / 12);
                int finalLabelW = Math.Max(scaled ? 76 : 110, usableW / 8);
                int valueW = (usableW - labelW * 2 - finalLabelW) / 3;
                DrawInfoCell(g, gridPen, head, x, y, labelW, infoH, "\uC791\uC5C5\uC790");
                DrawInfoCell(g, gridPen, font, x + labelW, y, valueW, infoH, userName);
                DrawInfoCell(g, gridPen, head, x + labelW + valueW, y, labelW, infoH, "\uC2DC\uD5D8\n\uC77C\uC790");
                DrawInfoCell(g, gridPen, font, x + labelW * 2 + valueW, y, valueW, infoH, evalDate);
                DrawInfoCell(g, gridPen, head, x + labelW * 2 + valueW * 2, y, finalLabelW, infoH, "\uCD5C\uC885\uC810\uC218");
                DrawInfoCell(g, gridPen, font, x + labelW * 2 + valueW * 2 + finalLabelW, y, valueW, infoH, "");
                y += infoH + Math.Max(5, padY);

                var objective = new List<int>();
                var subjective = new List<int>();
                for (int i = 0; i < questions.Count; i++)
                {
                    if (questions[i].Type == "\uC8FC\uAD00\uC2DD") subjective.Add(i);
                    else objective.Add(i);
                }
                int blockGap = Math.Max(8, usableW / 55);
                int blockCount = objective.Count > 12 ? 2 : 1;
                int blockW = blockCount == 2 ? (usableW - blockGap) / 2 : usableW;
                int leftCount = blockCount == 2 ? (objective.Count + 1) / 2 : objective.Count;
                int rightCount = objective.Count - leftCount;
                int objRows = Math.Max(leftCount, rightCount) + (objective.Count > 0 ? 1 : 0);
                int subjRows = subjective.Count + (subjective.Count > 0 ? 1 : 0);
                int gapY = subjective.Count > 0 && objective.Count > 0 ? Math.Max(10, padY * 2) : 0;
                int remain = Math.Max(1, bounds.Bottom - y - padY - gapY);
                int rowH = Math.Max(scaled ? 18 : 24, Math.Min(scaled ? 34 : 44, (int)(remain / Math.Max(1.0, objRows * 1.15 + subjRows * 3.2))));
                int subjRowH = Math.Max(scaled ? 70 : 86, rowH * 3);

                if (objective.Count > 0)
                {
                    DrawAnswerBlock(g, gridPen, head, font, questions, objective, 0, leftCount, x, y, blockW, rowH);
                    if (blockCount == 2)
                        DrawAnswerBlock(g, gridPen, head, font, questions, objective, leftCount, rightCount, x + blockW + blockGap, y, blockW, rowH);
                    y += objRows * rowH + gapY;
                }
                if (subjective.Count > 0)
                    DrawAnswerBlock(g, gridPen, head, font, questions, subjective, 0, subjective.Count, x, y, usableW, subjRowH);
            }
        }
        static void DrawInfoCell(Graphics g, Pen pen, Font font, int x, int y, int w, int h, string text)
        {
            var rect = new Rectangle(x, y, w, h);
            g.DrawRectangle(pen, rect);
            CenterText(g, text, font, Brushes.Black, rect);
        }
        static void DrawAnswerBlock(Graphics g, Pen pen, Font head, Font font, List<Question> questions, List<int> indices, int start, int count, int x, int y, int blockW, int rowH)
        {
            int noW = Math.Max(32, blockW / 8);
            int scoreW = Math.Max(38, blockW / 8);
            int checkW = Math.Max(42, blockW / 8);
            int ansW = blockW - noW - scoreW - checkW;
            using (var gray = new SolidBrush(Color.FromArgb(217, 217, 217)))
                g.FillRectangle(gray, new Rectangle(x, y, blockW, rowH));
            DrawAnswerRow(g, pen, head, font, x, y, noW, ansW, scoreW, checkW, rowH, "NO", "\uB2F5\uC548\uC9C0", "\uC810\uC218", "\uCC44\uC810", true);
            y += rowH;
            for (int i = 0; i < count; i++)
            {
                int qIndex = indices[start + i];
                Question q = questions[qIndex];
                DrawAnswerRow(g, pen, head, font, x, y, noW, ansW, scoreW, checkW, rowH, (qIndex + 1).ToString(), Preview.AnswerMark(q), q.Score.ToString("0.##"), "", false);
                y += rowH;
            }
        }
        static void DrawAnswerRow(Graphics g, Pen pen, Font head, Font font, int x, int y, int noW, int ansW, int scoreW, int checkW, int rowH, string no, string answer, string score, string check, bool header)
        {
            var noRect = new Rectangle(x, y, noW, rowH);
            var ansRect = new Rectangle(x + noW, y, ansW, rowH);
            var scoreRect = new Rectangle(x + noW + ansW, y, scoreW, rowH);
            var checkRect = new Rectangle(x + noW + ansW + scoreW, y, checkW, rowH);
            g.DrawRectangle(pen, noRect);
            g.DrawRectangle(pen, ansRect);
            g.DrawRectangle(pen, scoreRect);
            g.DrawRectangle(pen, checkRect);
            CenterText(g, no, header ? head : font, Brushes.Black, noRect);
            if (header)
                CenterText(g, answer, head, Brushes.Black, ansRect);
            else
                FitAnswerText(g, answer, font, Brushes.Black, ansRect);
            CenterText(g, score, header ? head : font, Brushes.Black, scoreRect);
            CenterText(g, check, header ? head : font, Brushes.Black, checkRect);
        }
        static void FitAnswerText(Graphics g, string text, Font baseFont, Brush brush, Rectangle rect)
        {
            Rectangle inner = new Rectangle(rect.X + 4, rect.Y + 2, rect.Width - 8, rect.Height - 4);
            using (var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, FormatFlags = 0, Trimming = StringTrimming.EllipsisCharacter })
            {
                for (float size = baseFont.Size; size >= 6f; size -= .5f)
                {
                    using (var f = new Font(baseFont.FontFamily, size, baseFont.Style))
                    {
                        SizeF measured = g.MeasureString(text ?? "", f, inner.Size, sf);
                        if (measured.Height <= inner.Height || size <= 6f)
                        {
                            g.DrawString(text ?? "", f, brush, inner, sf);
                            return;
                        }
                    }
                }
            }
        }
        static void CenterText(Graphics g, string text, Font font, Brush brush, Rectangle rect)
        {
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(text ?? "", font, brush, rect, sf);
        }
        void Status(string t, bool err)
        {
            if (status == null) return;
            string message = t == "LOADING" ? "문제은행 확인 중" : (t == "READY" ? "문제은행 연결됨" : (t == "ERROR" ? "문제은행 오류" : t));
            status.Text = message;
            status.ForeColor = err ? Color.FromArgb(230, 84, 84) : Color.FromArgb(73, 181, 137);
            if (secondaryStatus != null)
            {
                secondaryStatus.Text = message;
                secondaryStatus.ForeColor = status.ForeColor;
            }
        }
        static string DefaultWorkbook()
        {
            string saved = LoadLastWorkbook();
            if (File.Exists(saved)) return saved;
            foreach (string dir in new[] { Environment.CurrentDirectory, AppDomain.CurrentDomain.BaseDirectory, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) })
                if (Directory.Exists(dir))
                    foreach (string f in Directory.GetFiles(dir, "*.xlsm")) if (!Path.GetFileName(f).StartsWith("~$")) return f;
            return Path.Combine(Environment.CurrentDirectory, "OJT \uC2DC\uD5D8 \uBB38\uC81C.xlsm");
        }
        static string SettingsPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OJT_EXAM_MAKER", "last_workbook.txt");
        }
        static string LoadLastWorkbook()
        {
            try
            {
                string p = SettingsPath();
                return File.Exists(p) ? File.ReadAllText(p).Trim() : "";
            }
            catch { return ""; }
        }
        static void SaveLastWorkbook(string file)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(file)) return;
                string p = SettingsPath();
                Directory.CreateDirectory(Path.GetDirectoryName(p));
                File.WriteAllText(p, file.Trim());
            }
            catch { }
        }
    }

    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--workflow-self-test")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string workbook = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                using (var form = new MainForm()) form.RunWorkflowSelfTest(workbook);
                return;
            }
            if (args.Length > 0 && args[0] == "--ui-render")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string output = Path.GetFullPath(args.Length > 1 ? args[1] : "ui_render.png");
                bool workspace = args.Skip(2).Any(value => value == "workspace");
                bool small = args.Skip(2).Any(value => value == "small");
                using (var form = new MainForm())
                {
                    form.FormBorderStyle = FormBorderStyle.None;
                    form.ClientSize = small ? new Size(1280, 720) : new Size(1600, 900);
                    if (workspace) form.PrepareWorkspaceForTest();
                    form.Show();
                    Application.DoEvents();
                    form.PerformLayout();
                    System.Threading.Thread.Sleep(600);
                    form.Refresh();
                    Application.DoEvents();
                    using (var bitmap = new Bitmap(form.ClientSize.Width, form.ClientSize.Height))
                    {
                        form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.ClientSize));
                        Directory.CreateDirectory(Path.GetDirectoryName(output));
                        bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    form.Close();
                }
                return;
            }
            if (args.Length > 0 && args[0] == "--self-test")
            {
                string f = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                var banks = new Reader().Load(f);
                var all = banks.SelectMany(b => b.Questions).ToList();
                var noImageSample = all.FirstOrDefault(q => q.Text.IndexOf("\uC0DD\uC0B0 \uC77C\uC815\uC774 \uCD09\uBC15", StringComparison.Ordinal) >= 0);
                var imageSample = all.FirstOrDefault(q => q.Text.IndexOf("CZ\uACF5\uC815\uC740 \uC81C\uD488 \uD45C\uBA74", StringComparison.Ordinal) >= 0);
                int noImageCount = noImageSample == null ? -1 : noImageSample.Images.Count;
                int imageCount = imageSample == null ? -1 : imageSample.Images.Count;
                Console.WriteLine("banks={0} questions={1} images={2} no_image_sample={3} image_sample={4}", banks.Count, all.Count, all.Sum(q => q.Images.Count), noImageCount, imageCount);
                if (noImageSample != null && noImageCount != 0)
                    throw new InvalidOperationException("A regular question received an unrelated image.");
                if (imageSample != null && imageCount == 0)
                    throw new InvalidOperationException("An image question lost its source image.");
                return;
            }
            if (args.Length > 0 && args[0] == "--render-test")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string f = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                string output = Path.GetFullPath(args.Length > 2 ? args[2] : "render_test.png");
                var banks = new Reader().Load(f);
                var bank = banks.FirstOrDefault(b => b.Questions.Any(q => q.Text.IndexOf("CZ\uACF5\uC815\uC740 \uC81C\uD488 \uD45C\uBA74", StringComparison.Ordinal) >= 0));
                if (bank == null) throw new InvalidOperationException("CZ image sample was not found.");
                var regular = bank.Questions.First(q => q.Text.IndexOf("\uC0DD\uC0B0 \uC77C\uC815\uC774 \uCD09\uBC15", StringComparison.Ordinal) >= 0);
                var imageQuestion = bank.Questions.First(q => q.Text.IndexOf("CZ\uACF5\uC815\uC740 \uC81C\uD488 \uD45C\uBA74", StringComparison.Ordinal) >= 0);
                var subjective = bank.Questions.FirstOrDefault(q => q.Type == "\uC8FC\uAD00\uC2DD") ?? bank.Questions.Last();
                var multiImage = banks.SelectMany(b => b.Questions).FirstOrDefault(q => q.Images.Count == 4 && q.Text.IndexOf("03BGASEMBLDN", StringComparison.Ordinal) >= 0);
                var longQuestion = banks.SelectMany(b => b.Questions)
                    .Where(q => q.Images.Count == 0 && q.Type != "\uC8FC\uAD00\uC2DD")
                    .OrderByDescending(q => (q.Text ?? "").Split('\n').Select(line => line.Length).DefaultIfEmpty(0).Max())
                    .FirstOrDefault();
                var doc = new ExamDoc { Bank = bank, UserName = "TEST", EvalDate = DateTime.Today.ToString("yyyy.MM.dd") };
                doc.Questions.Add(regular);
                doc.Questions.Add(imageQuestion);
                if (multiImage != null) doc.Questions.Add(multiImage);
                if (longQuestion != null) doc.Questions.Add(longQuestion);
                doc.Questions.Add(subjective);
                using (var bitmap = new Bitmap(827, 1169))
                using (var graphics = Graphics.FromImage(bitmap))
                using (var preview = new Preview { Doc = doc })
                {
                    graphics.Clear(Color.White);
                    preview.DrawPrint(graphics, new Rectangle(18, 18, 791, 1133), 0, false);
                    Directory.CreateDirectory(Path.GetDirectoryName(output));
                    bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                }
                Console.WriteLine(output);
                return;
            }
            if (args.Length > 0 && args[0] == "--screen-preview-render-test")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string f = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                string output = Path.GetFullPath(args.Length > 2 ? args[2] : "screen_preview_test.png");
                var banks = new Reader().Load(f);
                var bank = banks.FirstOrDefault(b => b.Questions.Count(q => q.Type == "VDA") >= 5);
                if (bank == null) throw new InvalidOperationException("A screen-preview test bank was not found.");
                var testDoc = new ExamDoc { Bank = bank, UserName = "TEST", EvalDate = DateTime.Today.ToString("yyyy.MM.dd") };
                testDoc.Questions.AddRange(bank.Questions.Where(q => q.Type == "VDA").Take(5));
                using (var screen = new Preview { Doc = testDoc, Size = new Size(646, 900) })
                using (var bitmap = new Bitmap(646, 900))
                {
                    screen.CreateControl();
                    screen.DrawToBitmap(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                    Directory.CreateDirectory(Path.GetDirectoryName(output));
                    bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                }
                Console.WriteLine(output);
                return;
            }
            if (args.Length > 0 && args[0] == "--request-render-test")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string f = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                string output = Path.GetFullPath(args.Length > 2 ? args[2] : "request_render_test.png");
                var banks = new Reader().Load(f);
                var all = banks.SelectMany(b => b.Questions).ToList();
                var hanging = all.FirstOrDefault(q => q.Text.IndexOf("V-PRESS \uC131\uD615 \u524D \uC5D0\uCE6D \uC57D\uD488", StringComparison.Ordinal) >= 0);
                var multiImage = all.FirstOrDefault(q => q.Images.Count == 4 && q.Text.IndexOf("03BGASEMBLDN", StringComparison.Ordinal) >= 0);
                var spaced = all.FirstOrDefault(q => q.Images.Count == 0 && q.Text.Contains("\n\n") && q.Text.IndexOf("[VDA 6.3]", StringComparison.Ordinal) >= 0);
                var fixedScore = all.FirstOrDefault(q => q.Text.IndexOf("Lay-up PM\uD6C4 \uBAA9\uC2DC\uB85C \uC774\uBB3C", StringComparison.Ordinal) >= 0);
                var selected = new[] { hanging, multiImage, spaced, fixedScore }.Where(q => q != null).ToList();
                if (selected.Count < 4) throw new InvalidOperationException("Request verification samples were not found.");
                var bank = banks.First(b => b.Questions.Contains(hanging));
                var doc = new ExamDoc { Bank = bank, UserName = "TEST", EvalDate = DateTime.Today.ToString("yyyy.MM.dd") };
                doc.Questions.AddRange(selected);
                using (var bitmap = new Bitmap(827, 1169))
                using (var graphics = Graphics.FromImage(bitmap))
                using (var preview = new Preview { Doc = doc })
                {
                    graphics.Clear(Color.White);
                    preview.DrawPrint(graphics, new Rectangle(18, 18, 791, 1133), 0, false);
                    Directory.CreateDirectory(Path.GetDirectoryName(output));
                    bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                }
                Console.WriteLine(output);
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
        static string MainFormDefaultWorkbook()
        {
            foreach (string f in Directory.GetFiles(Environment.CurrentDirectory, "*.xlsm")) if (!Path.GetFileName(f).StartsWith("~$")) return f;
            return Path.Combine(Environment.CurrentDirectory, "OJT \uC2DC\uD5D8 \uBB38\uC81C.xlsm");
        }
    }
}

