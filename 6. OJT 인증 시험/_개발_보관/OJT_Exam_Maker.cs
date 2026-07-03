using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace OJT_Exam_Maker
{
    internal sealed class Question
    {
        public string SourceSheet = "";
        public string SourceNo = "";
        public string ExamType = "";
        public string Text = "";
        public string Answer = "";
        public decimal Score;
        public readonly List<Image> Images = new List<Image>();
    }

    internal sealed class QuestionBank
    {
        public string Name = "";
        public string Department = "";
        public string Evaluator = "";
        public string JobName = "";
        public string Revision = "";
        public string IssueDate = "";
        public string ProductType = "";
        public readonly List<Question> Questions = new List<Question>();
        public int Common { get { return Questions.Count(q => q.ExamType == "공통"); } }
        public int Choice { get { return Questions.Count(q => q.ExamType == "객관식"); } }
        public int Subjective { get { return Questions.Count(q => q.ExamType == "주관식"); } }
    }

    internal sealed class ExamDocument
    {
        public string ProcessName = "";
        public string Department = "후가공";
        public string Evaluator = "김용준";
        public string JobName = "";
        public string Revision = "0";
        public string IssueDate = "2024.12.09";
        public string ProductType = "";
        public DateTime CreatedAt = DateTime.Now;
        public bool ShowAnswers;
        public readonly List<Question> Questions = new List<Question>();
    }

    internal sealed class SheetInfo
    {
        public string Name = "";
        public string Path = "";
    }

    internal sealed class HeaderInfo
    {
        public int Row;
        public int NoCol;
        public int TypeCol;
        public int QuestionCol;
        public int AnswerCol;
        public int ScoreCol;
        public readonly List<int> ChoiceCols = new List<int>();
    }

    internal sealed class SettingInfo
    {
        public string Department = "";
        public string Evaluator = "";
        public string JobName = "";
        public string Revision = "";
        public string IssueDate = "";
        public string ProductType = "";
    }

    internal sealed class XlsmQuestionReader
    {
        private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public List<QuestionBank> Load(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = LoadSharedStrings(zip);
                var sheets = LoadSheets(zip);
                var settings = LoadSettings(zip, sheets, shared);
                var banks = new List<QuestionBank>();

                foreach (var sheet in sheets)
                {
                    if (sheet.Name == "시험 SETTING" || sheet.Name == "시험지" || sheet.Name == "답안지")
                        continue;

                    var cells = LoadCells(zip, sheet.Path, shared);
                    var header = FindHeader2(cells);
                    if (header == null)
                        continue;

                    var images = LoadImagesByRow(zip, sheet.Path);
                    var bank = new QuestionBank { Name = sheet.Name.Trim() };
                    ApplySettings(bank, settings);
                    int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
                    for (int row = header.Row + 1; row <= maxRow; row++)
                    {
                        string text = BuildQuestionText(cells, row, header);
                        string no = Clean(Get(cells, row, header.NoCol));
                        if (text.Length == 0 || no.Length == 0)
                            continue;

                        string type = Clean(Get(cells, row, header.TypeCol));
                        if (type != "공통" && type != "객관식" && type != "주관식")
                            continue;

                        decimal score = 0;
                        decimal.TryParse(Clean(Get(cells, row, header.ScoreCol)), out score);
                        var q = new Question
                        {
                            SourceSheet = sheet.Name,
                            SourceNo = no,
                            ExamType = type,
                            Text = NormalizeQuestionText(text),
                            Answer = Clean(Get(cells, row, header.AnswerCol)),
                            Score = score
                        };
                        if (images.ContainsKey(row))
                            q.Images.AddRange(images[row]);
                        bank.Questions.Add(q);
                    }

                    if (bank.Questions.Count > 0)
                        banks.Add(bank);
                }
                return banks.OrderBy(b => b.Name).ToList();
            }
        }

        private static Dictionary<string, SettingInfo> LoadSettings(ZipArchive zip, List<SheetInfo> sheets, List<string> shared)
        {
            var result = new Dictionary<string, SettingInfo>(StringComparer.OrdinalIgnoreCase);
            var settingSheet = sheets.FirstOrDefault(s => s.Name.Trim() == "\uC2DC\uD5D8 SETTING");
            if (settingSheet == null)
                return result;
            var cells = LoadCells(zip, settingSheet.Path, shared);
            int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
            for (int row = 3; row <= maxRow; row++)
            {
                string key = Clean(Get(cells, row, 11)).Trim();
                if (key.Length == 0)
                    continue;
                result[key] = new SettingInfo
                {
                    Department = Clean(Get(cells, row, 12)),
                    Evaluator = Clean(Get(cells, row, 13)),
                    JobName = Clean(Get(cells, row, 14)),
                    Revision = Clean(Get(cells, row, 15)),
                    IssueDate = Clean(Get(cells, row, 16)),
                    ProductType = Clean(Get(cells, row, 17))
                };
            }
            return result;
        }

        private static void ApplySettings(QuestionBank bank, Dictionary<string, SettingInfo> settings)
        {
            SettingInfo info;
            if (!settings.TryGetValue(bank.Name.Trim(), out info))
                settings.TryGetValue(bank.Name.TrimEnd(), out info);
            if (info != null)
            {
                bank.Department = info.Department;
                bank.Evaluator = info.Evaluator;
                bank.JobName = info.JobName;
                bank.Revision = info.Revision;
                bank.IssueDate = info.IssueDate;
                bank.ProductType = info.ProductType;
            }
            if (bank.Department.Length == 0) bank.Department = "\uD6C4\uAC00\uACF5";
            if (bank.Evaluator.Length == 0) bank.Evaluator = "\uAE40\uC6A9\uC900";
            if (bank.JobName.Length == 0)
                bank.JobName = bank.Name.Replace(" \uC77C\uBC18\uC6A9", "").Replace(" \uC804\uC7A5\uC6A9", "").Trim();
            if (bank.Revision.Length == 0) bank.Revision = "0";
            if (bank.IssueDate.Length == 0) bank.IssueDate = "2024.12.09";
            if (bank.ProductType.Length == 0)
            {
                if (bank.Name.Contains("\uC804\uC7A5\uC6A9")) bank.ProductType = "\uC804\uC7A5\uC6A9";
                else if (bank.Name.Contains("\uC77C\uBC18\uC6A9")) bank.ProductType = "\uC77C\uBC18\uC6A9";
            }
        }

        private static string BuildQuestionText(Dictionary<Tuple<int, int>, string> cells, int row, HeaderInfo header)
        {
            string text = Clean(Get(cells, row, header.QuestionCol));
            if (header.ChoiceCols.Count == 0)
                return text;
            string[] marks = { "\u2460", "\u2461", "\u2462", "\u2463", "\u2464", "\u2465" };
            var choices = new List<string>();
            for (int i = 0; i < header.ChoiceCols.Count && i < marks.Length; i++)
            {
                string choice = Clean(Get(cells, row, header.ChoiceCols[i]));
                if (choice.Length > 0)
                    choices.Add(marks[i] + " " + choice);
            }
            return choices.Count == 0 ? text : text + "\n\n" + string.Join("\n", choices);
        }

        private static List<string> LoadSharedStrings(ZipArchive zip)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            var result = new List<string>();
            if (entry == null)
                return result;
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var si in doc.Descendants(MainNs + "si"))
                    result.Add(string.Concat(si.Descendants(MainNs + "t").Select(t => (string)t)));
            }
            return result;
        }

        private static List<SheetInfo> LoadSheets(ZipArchive zip)
        {
            var rels = new Dictionary<string, string>();
            using (var s = zip.GetEntry("xl/_rels/workbook.xml.rels").Open())
            {
                var doc = XDocument.Load(s);
                foreach (var r in doc.Root.Elements(PackageRelNs + "Relationship"))
                {
                    var id = (string)r.Attribute("Id");
                    var target = ResolvePart("xl/workbook.xml", (string)r.Attribute("Target"));
                    rels[id] = target;
                }
            }

            var list = new List<SheetInfo>();
            using (var s = zip.GetEntry("xl/workbook.xml").Open())
            {
                var doc = XDocument.Load(s);
                foreach (var sh in doc.Descendants(MainNs + "sheet"))
                {
                    string id = (string)sh.Attribute(RelNs + "id");
                    if (id != null && rels.ContainsKey(id))
                        list.Add(new SheetInfo { Name = (string)sh.Attribute("name") ?? "", Path = rels[id] });
                }
            }
            return list;
        }

        private static Dictionary<Tuple<int, int>, string> LoadCells(ZipArchive zip, string sheetPath, List<string> shared)
        {
            var result = new Dictionary<Tuple<int, int>, string>();
            var entry = zip.GetEntry(sheetPath);
            if (entry == null)
                return result;

            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var c in doc.Descendants(MainNs + "c"))
                {
                    string r = (string)c.Attribute("r");
                    int row, col;
                    ParseRef(r, out row, out col);
                    string type = (string)c.Attribute("t");
                    string value = "";
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
                    result[Tuple.Create(row, col)] = value;
                }
            }
            return result;
        }

        private static Dictionary<int, List<Image>> LoadImagesByRow(ZipArchive zip, string sheetPath)
        {
            var result = new Dictionary<int, List<Image>>();
            string relPath = PathToRels(sheetPath);
            var sheetRels = zip.GetEntry(relPath);
            if (sheetRels == null)
                return result;

            string drawingPath = null;
            using (var s = sheetRels.Open())
            {
                var doc = XDocument.Load(s);
                var drawingRel = doc.Root.Elements(PackageRelNs + "Relationship")
                    .FirstOrDefault(r => ((string)r.Attribute("Type") ?? "").EndsWith("/drawing"));
                if (drawingRel != null)
                    drawingPath = ResolvePart(sheetPath, (string)drawingRel.Attribute("Target"));
            }
            if (drawingPath == null || zip.GetEntry(drawingPath) == null)
                return result;

            var drawingRels = new Dictionary<string, string>();
            var drawingRelsEntry = zip.GetEntry(PathToRels(drawingPath));
            if (drawingRelsEntry != null)
            {
                using (var s = drawingRelsEntry.Open())
                {
                    var doc = XDocument.Load(s);
                    foreach (var r in doc.Root.Elements(PackageRelNs + "Relationship"))
                        drawingRels[(string)r.Attribute("Id")] = ResolvePart(drawingPath, (string)r.Attribute("Target"));
                }
            }

            using (var s = zip.GetEntry(drawingPath).Open())
            {
                var doc = XDocument.Load(s);
                foreach (var anchor in doc.Root.Elements())
                {
                    var from = anchor.Element(DrawingNs + "from");
                    var blip = anchor.Descendants(A + "blip").FirstOrDefault();
                    if (from == null || blip == null)
                        continue;
                    int row = ToInt((string)from.Element(DrawingNs + "row")) + 1;
                    string rid = (string)blip.Attribute(RelNs + "embed");
                    if (rid == null || !drawingRels.ContainsKey(rid))
                        continue;
                    var imgEntry = zip.GetEntry(drawingRels[rid]);
                    if (imgEntry == null)
                        continue;
                    using (var ims = imgEntry.Open())
                    using (var ms = new MemoryStream())
                    {
                        ims.CopyTo(ms);
                        ms.Position = 0;
                        Image img = Image.FromStream(ms);
                        if (!result.ContainsKey(row))
                            result[row] = new List<Image>();
                        result[row].Add(new Bitmap(img));
                    }
                }
            }
            return result;
        }

        private static HeaderInfo FindHeader(Dictionary<Tuple<int, int>, string> cells)
        {
            int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
            int maxCol = cells.Keys.Select(k => k.Item2).DefaultIfEmpty(0).Max();
            for (int row = 1; row <= Math.Min(12, maxRow); row++)
            {
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= maxCol; col++)
                {
                    string h = Normalize(Get(cells, row, col));
                    if (h.Length > 0 && !headers.ContainsKey(h))
                        headers[h] = col;
                }
                if (headers.ContainsKey("no") && headers.ContainsKey("문제유형") && headers.ContainsKey("문제") && headers.ContainsKey("답안"))
                {
                    return new HeaderInfo
                    {
                        Row = row,
                        NoCol = headers["no"],
                        TypeCol = headers["문제유형"],
                        QuestionCol = headers["문제"],
                        AnswerCol = headers["답안"],
                        ScoreCol = headers.ContainsKey("점수") ? headers["점수"] : 0
                    };
                }
            }
            return null;
        }

        private static string Get(Dictionary<Tuple<int, int>, string> cells, int row, int col)
        {
            if (row <= 0 || col <= 0)
                return "";
            string value;
            return cells.TryGetValue(Tuple.Create(row, col), out value) ? value : "";
        }

        private static HeaderInfo FindHeader2(Dictionary<Tuple<int, int>, string> cells)
        {
            int maxRow = cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
            int maxCol = cells.Keys.Select(k => k.Item2).DefaultIfEmpty(0).Max();
            for (int row = 1; row <= Math.Min(12, maxRow); row++)
            {
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= maxCol; col++)
                {
                    string h = Normalize(Get(cells, row, col));
                    if (h.Length > 0 && !headers.ContainsKey(h))
                        headers[h] = col;
                }

                string typeKey = FirstHeaderKey(headers, "\uBB38\uC81C\uC720\uD615", "臾몄젣?좏삎");
                string questionKey = FirstHeaderKey(headers, "\uBB38\uC81C", "臾몄젣");
                string answerKey = FirstHeaderKey(headers, "\uB2F5\uC548", "\uC815\uB2F5", "?듭븞");
                string scoreKey = FirstHeaderKey(headers, "\uC810\uC218", "?먯닔");
                if (headers.ContainsKey("no") && typeKey != null && questionKey != null && answerKey != null)
                {
                    var info = new HeaderInfo
                    {
                        Row = row,
                        NoCol = headers["no"],
                        TypeCol = headers[typeKey],
                        QuestionCol = headers[questionKey],
                        AnswerCol = headers[answerKey],
                        ScoreCol = scoreKey != null ? headers[scoreKey] : 0
                    };
                    for (int n = 1; n <= 6; n++)
                    {
                        string key = n.ToString();
                        if (headers.ContainsKey(key))
                            info.ChoiceCols.Add(headers[key]);
                    }
                    return info;
                }
            }
            return null;
        }

        private static string FirstHeaderKey(Dictionary<string, int> headers, params string[] keys)
        {
            foreach (string key in keys)
                if (headers.ContainsKey(key))
                    return key;
            return null;
        }

        private static string Clean(string value)
        {
            if (value == null)
                return "";
            value = value.Replace("\r\n", "\n").Replace("\r", "\n");
            return Regex.Replace(value, @"[ \t]+\n", "\n").Trim();
        }

        private static string Normalize(string value)
        {
            return Regex.Replace(Clean(value), @"\s+", "").ToLowerInvariant();
        }

        private static string NormalizeQuestionText(string value)
        {
            return Clean(value).Replace(" ( )", "( )").Replace("(  )", "( )");
        }

        private static int ToInt(string value)
        {
            int n;
            return int.TryParse(value, out n) ? n : 0;
        }

        private static void ParseRef(string reference, out int row, out int col)
        {
            row = 0;
            col = 0;
            if (string.IsNullOrEmpty(reference))
                return;
            foreach (char ch in reference)
            {
                if (char.IsLetter(ch))
                    col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                else if (char.IsDigit(ch))
                    row = row * 10 + (ch - '0');
            }
        }

        private static string ResolvePart(string basePart, string target)
        {
            if (string.IsNullOrEmpty(target))
                return "";
            target = target.Replace('\\', '/');
            if (target.StartsWith("/"))
                return target.TrimStart('/');
            var parts = basePart.Split('/').ToList();
            parts.RemoveAt(parts.Count - 1);
            foreach (string piece in target.Split('/'))
            {
                if (piece == "." || piece.Length == 0)
                    continue;
                if (piece == ".." && parts.Count > 0)
                    parts.RemoveAt(parts.Count - 1);
                else if (piece != "..")
                    parts.Add(piece);
            }
            return string.Join("/", parts);
        }

        private static string PathToRels(string part)
        {
            int slash = part.LastIndexOf('/');
            string dir = slash >= 0 ? part.Substring(0, slash + 1) : "";
            string file = slash >= 0 ? part.Substring(slash + 1) : part;
            return dir + "_rels/" + file + ".rels";
        }
    }

    internal sealed class PreviewCanvas : Control
    {
        public ExamDocument Document;
        public int PageIndex;
        private readonly List<int> pageStarts = new List<int>();

        public PreviewCanvas()
        {
            DoubleBuffered = true;
            BackColor = Color.FromArgb(7, 16, 29);
            ResizeRedraw = true;
        }

        public int PageCount
        {
            get
            {
                RebuildPages();
                return Math.Max(1, pageStarts.Count);
            }
        }

        public void MovePage(int delta)
        {
            RebuildPages();
            PageIndex = Math.Max(0, Math.Min(PageIndex + delta, PageCount - 1));
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            if (Document == null)
            {
                using (var brush = new SolidBrush(Color.FromArgb(188, 209, 255)))
                using (var font = new Font("맑은 고딕", 13, FontStyle.Bold))
                    e.Graphics.DrawString("문항 조건을 선택하고 [랜덤 미리보기]를 누르세요.", font, brush, new PointF(30, 32));
                return;
            }

            RebuildPages();
            var page = GetPageRect(ClientRectangle);
            using (var paper = new SolidBrush(Color.White))
            using (var border = new Pen(Color.FromArgb(80, 95, 115)))
            {
                e.Graphics.FillRectangle(paper, page);
                e.Graphics.DrawRectangle(border, page);
            }
            DrawExamPage(e.Graphics, page, PageIndex, Document.ShowAnswers);
        }

        public void DrawForPrint(Graphics g, Rectangle bounds, int pageIndex, bool showAnswers)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            DrawExamPage(g, bounds, pageIndex, showAnswers);
        }

        private Rectangle GetPageRect(Rectangle bounds)
        {
            float ratio = 1122f / 793f;
            int w = bounds.Width - 28;
            int h = (int)(w / ratio);
            if (h > bounds.Height - 28)
            {
                h = bounds.Height - 28;
                w = (int)(h * ratio);
            }
            return new Rectangle(bounds.X + (bounds.Width - w) / 2, bounds.Y + 14, w, h);
        }

        private void RebuildPages()
        {
            pageStarts.Clear();
            if (Document == null || Document.Questions.Count == 0)
                return;
            pageStarts.Add(0);
            int used = 0;
            const int pageLimit = 600;
            for (int i = 0; i < Document.Questions.Count; i++)
            {
                int h = EstimateQuestionHeight(Document.Questions[i]);
                if (i > 0 && used + h > pageLimit)
                {
                    pageStarts.Add(i);
                    used = 0;
                }
                used += h;
            }
            if (PageIndex >= pageStarts.Count)
                PageIndex = pageStarts.Count - 1;
        }

        private int LastQuestionIndexOnPage(int pageIndex)
        {
            if (pageIndex + 1 < pageStarts.Count)
                return pageStarts[pageIndex + 1];
            return Document.Questions.Count;
        }

        private void DrawExamPage(Graphics g, Rectangle page, int pageIndex, bool showAnswers)
        {
            float sx = page.Width / 1122f;
            float sy = page.Height / 793f;
            Func<float, int> X = v => page.Left + (int)(v * sx);
            Func<float, int> Y = v => page.Top + (int)(v * sy);
            Func<float, int> W = v => Math.Max(1, (int)(v * sx));
            Func<float, int> H = v => Math.Max(1, (int)(v * sy));

            using (var pen = new Pen(Color.Black, W(1.2f)))
            using (var thin = new Pen(Color.Black, W(.8f)))
            using (var text = new SolidBrush(Color.Black))
            using (var title = new Font("맑은 고딕", Math.Max(10, W(18)), FontStyle.Bold))
            using (var font = new Font("맑은 고딕", Math.Max(8, W(10)), FontStyle.Regular))
            using (var bold = new Font("맑은 고딕", Math.Max(8, W(10)), FontStyle.Bold))
            {
                DrawTemplateHeader(g, pen, text, title, font, bold, X, Y, W, H);

                int y = Y(194);
                int start = pageStarts.Count == 0 ? 0 : pageStarts[Math.Min(pageIndex, pageStarts.Count - 1)];
                int end = LastQuestionIndexOnPage(Math.Min(pageIndex, PageCount - 1));

                DrawCell(g, "No", bold, new Rectangle(X(22), y, W(48), H(24)), true);
                DrawCell(g, "문제", bold, new Rectangle(X(70), y, W(1030), H(24)), true);
                y += H(24);

                for (int i = start; i < end; i++)
                {
                    var q = Document.Questions[i];
                    int rowH = H(EstimateQuestionHeight(q));
                    DrawCell(g, (i + 1).ToString(), bold, new Rectangle(X(22), y, W(48), rowH), true);
                    DrawQuestionCell(g, q, font, bold, text, new Rectangle(X(70), y, W(1030), rowH), showAnswers);
                    g.DrawRectangle(thin, new Rectangle(X(22), y, W(1078), rowH));
                    y += rowH;
                }
            }
        }

        private void DrawTemplateHeader(Graphics g, Pen pen, Brush text, Font title, Font font, Font bold, Func<float, int> X, Func<float, int> Y, Func<float, int> W, Func<float, int> H)
        {
            DrawCell(g, "부서명", bold, new Rectangle(X(22), Y(8), W(92), H(24)), true);
            DrawCell(g, Document.Department, font, new Rectangle(X(114), Y(8), W(145), H(24)), true);
            DrawCell(g, "교육 평가서", title, new Rectangle(X(259), Y(8), W(615), H(94)), true);
            DrawCell(g, "기 안", bold, new Rectangle(X(874), Y(8), W(84), H(24)), true);
            DrawCell(g, "심의", bold, new Rectangle(X(958), Y(8), W(84), H(24)), true);
            DrawCell(g, "결정", bold, new Rectangle(X(1042), Y(8), W(58), H(24)), true);
            DrawCell(g, "평가자", bold, new Rectangle(X(22), Y(32), W(92), H(38)), true);
            DrawCell(g, Document.Evaluator, font, new Rectangle(X(114), Y(32), W(145), H(38)), true);
            DrawCell(g, "", font, new Rectangle(X(874), Y(32), W(84), H(46)), true);
            DrawCell(g, "", font, new Rectangle(X(958), Y(32), W(84), H(46)), true);
            DrawCell(g, "", font, new Rectangle(X(1042), Y(32), W(58), H(46)), true);
            DrawCell(g, "평가 일시", bold, new Rectangle(X(22), Y(70), W(92), H(32)), true);
            DrawCell(g, "", font, new Rectangle(X(114), Y(70), W(145), H(32)), true);
            DrawCell(g, "/", font, new Rectangle(X(874), Y(78), W(84), H(24)), true);
            DrawCell(g, "/", font, new Rectangle(X(958), Y(78), W(84), H(24)), true);
            DrawCell(g, "/", font, new Rectangle(X(1042), Y(78), W(58), H(24)), true);

            DrawCell(g, "■ 직 무 명", bold, new Rectangle(X(22), Y(114), W(230), H(24)), true);
            DrawCell(g, Document.JobName, font, new Rectangle(X(252), Y(114), W(460), H(24)), true);
            DrawCell(g, "■ 성 명", bold, new Rectangle(X(712), Y(114), W(95), H(24)), true);
            DrawCell(g, "", font, new Rectangle(X(807), Y(114), W(293), H(24)), true);
            DrawCell(g, "평가 방법", bold, new Rectangle(X(22), Y(138), W(230), H(24)), true);
            DrawCell(g, "시험 평가, 결과 보고서, 직무 평가, 기타 방법 (        )", font, new Rectangle(X(252), Y(138), W(848), H(24)), true);
            DrawCell(g, "개정 차수", bold, new Rectangle(X(22), Y(162), W(230), H(24)), true);
            DrawCell(g, Document.Revision, font, new Rectangle(X(252), Y(162), W(158), H(24)), true);
            DrawCell(g, "제정일", bold, new Rectangle(X(410), Y(162), W(158), H(24)), true);
            DrawCell(g, Document.IssueDate, font, new Rectangle(X(568), Y(162), W(158), H(24)), true);
            DrawCell(g, "유 형", bold, new Rectangle(X(726), Y(162), W(158), H(24)), true);
            DrawCell(g, Document.ProductType, font, new Rectangle(X(884), Y(162), W(216), H(24)), true);
            g.DrawRectangle(pen, new Rectangle(X(22), Y(8), W(1078), H(178)));
        }

        private static int EstimateQuestionHeight(Question q)
        {
            int lines = Math.Max(1, q.Text.Count(ch => ch == '\n') + 1);
            int textLines = Math.Max(lines, q.Text.Length / 95 + 1);
            int image = q.Images.Count == 0 ? 0 : Math.Min(120, q.Images.Sum(img => Math.Min(85, img.Height * 140 / Math.Max(1, img.Width))));
            int answer = q.ExamType == "주관식" ? 34 : 0;
            return Math.Max(q.Images.Count > 0 ? 105 : 42, textLines * 15 + image + answer + 14);
        }

        private static void DrawCell(Graphics g, string value, Font font, Rectangle rect, bool center)
        {
            using (var pen = new Pen(Color.Black, 1))
            {
                g.DrawRectangle(pen, rect);
            }
            var sf = new StringFormat
            {
                Alignment = center ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.LineLimit
            };
            var padded = new Rectangle(rect.X + 5, rect.Y + 2, rect.Width - 10, rect.Height - 4);
            using (var brush = new SolidBrush(Color.Black))
                g.DrawString(value ?? "", font, brush, padded, sf);
        }

        private static void DrawQuestionCell(Graphics g, Question q, Font font, Font bold, Brush brush, Rectangle rect, bool showAnswers)
        {
            using (var pen = new Pen(Color.Black, 1))
                g.DrawRectangle(pen, rect);
            string question = FormatQuestionWithScore(q);
            var textRect = new Rectangle(rect.X + 8, rect.Y + 7, rect.Width - 16, rect.Height - 14);
            if (q.Images.Count > 0)
            {
                int imgY = rect.Bottom - 8;
                foreach (var img in q.Images.Take(3).Reverse<Image>())
                {
                    int maxW = Math.Min(220, rect.Width - 20);
                    int imgW = maxW;
                    int imgH = Math.Max(35, img.Height * imgW / Math.Max(1, img.Width));
                    if (imgH > 85)
                    {
                        imgH = 85;
                        imgW = Math.Max(35, img.Width * imgH / Math.Max(1, img.Height));
                    }
                    imgY -= imgH;
                    g.DrawImage(img, new Rectangle(rect.X + 10, imgY, imgW, imgH));
                    imgY -= 6;
                }
                textRect.Height = Math.Max(28, imgY - textRect.Y);
            }
            g.DrawString(question, font, brush, textRect);
            if (q.ExamType == "주관식")
            {
                int lineY = rect.Bottom - 30;
                using (var line = new Pen(Color.Black, 1))
                    g.DrawLine(line, rect.X + 16, lineY, rect.Right - 16, lineY);
            }
            if (showAnswers)
            {
                string answer = ApplyAnswer(q);
                using (var red = new SolidBrush(Color.FromArgb(210, 0, 0)))
                    g.DrawString(answer, bold, red, new Rectangle(rect.Right - 90, rect.Y + 6, 80, 24));
            }
        }

        private static void DrawAnswerCell(Graphics g, Question q, Font font, Font bold, Brush brush, Rectangle rect, bool showAnswers)
        {
            using (var pen = new Pen(Color.Black, 1))
                g.DrawRectangle(pen, rect);
            if (q.ExamType == "주관식")
            {
                int lineY = rect.Y + rect.Height / 2;
                using (var pen = new Pen(Color.Black, 1))
                {
                    g.DrawLine(pen, rect.X + 10, lineY, rect.Right - 10, lineY);
                    g.DrawLine(pen, rect.X + 10, lineY + 28, rect.Right - 10, lineY + 28);
                }
            }
            else
            {
                using (var pen = new Pen(Color.Black, 1))
                {
                    int cx = rect.X + 18;
                    int cy = rect.Y + rect.Height / 2;
                    g.DrawEllipse(pen, cx, cy - 8, 16, 16);
                    g.DrawString("번", font, brush, rect.X + 40, cy - 10);
                }
            }
            if (showAnswers)
            {
                string answer = ApplyAnswer(q);
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(answer, bold, brush, rect, sf);
            }
        }

        private static void DrawCentered(Graphics g, string value, Font font, Brush brush, Rectangle rect)
        {
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(value, font, brush, rect, sf);
        }

        private static string FormatQuestionWithScore(Question q)
        {
            return "[" + q.ExamType + " / " + q.Score.ToString("0.##") + "점] " + q.Text;
        }

        public static string ApplyAnswer(Question q)
        {
            if (q.ExamType == "객관식" || q.ExamType == "공통")
            {
                int n = ExtractChoiceNumber(q.Answer, q.Text);
                if (n >= 1 && n <= 10)
                    return new[] { "", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩" }[n];
            }
            return q.Answer;
        }

        private static int ExtractChoiceNumber(string answer, string question)
        {
            var m = Regex.Match(answer ?? "", @"[①②③④⑤⑥⑦⑧⑨⑩1-9]");
            if (m.Success)
            {
                string circles = "①②③④⑤⑥⑦⑧⑨⑩";
                int idx = circles.IndexOf(m.Value, StringComparison.Ordinal);
                if (idx >= 0)
                    return idx + 1;
                return int.Parse(m.Value);
            }

            if (!string.IsNullOrWhiteSpace(answer) && !string.IsNullOrWhiteSpace(question))
            {
                var pattern = @"(?:^|\n)\s*([①②③④⑤⑥⑦⑧⑨⑩1-9])[\).\s]*(.+?)(?=\n\s*[①②③④⑤⑥⑦⑧⑨⑩1-9][\).\s]|$)";
                foreach (Match opt in Regex.Matches(question, pattern, RegexOptions.Singleline))
                {
                    string optionText = Regex.Replace(opt.Groups[2].Value, @"\s+", "");
                    string answerText = Regex.Replace(answer, @"\s+", "");
                    if (optionText.Contains(answerText) || answerText.Contains(optionText))
                    {
                        string mark = opt.Groups[1].Value;
                        int idx = "①②③④⑤⑥⑦⑧⑨⑩".IndexOf(mark, StringComparison.Ordinal);
                        return idx >= 0 ? idx + 1 : int.Parse(mark);
                    }
                }
            }
            return 0;
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly XlsmQuestionReader reader = new XlsmQuestionReader();
        private readonly Random random = new Random();
        private readonly List<QuestionBank> banks = new List<QuestionBank>();
        private TextBox pathBox;
        private DataGridView bankGrid;
        private NumericUpDown commonCount, choiceCount, subjectiveCount, commonScore, choiceScore, subjectiveScore, targetScore;
        private Label totalLabel, statusLabel, selectedLabel, pageLabel;
        private PreviewCanvas preview;
        private CheckBox answerCheck;
        private ExamDocument currentDoc;

        public MainForm()
        {
            Text = "OJT EXAM MAKER";
            Width = 1880;
            Height = 900;
            MinimumSize = new Size(1280, 720);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("맑은 고딕", 9F);
            BackColor = Color.FromArgb(11, 18, 32);
            BuildUi();
            pathBox.Text = DefaultWorkbookPath2();
            LoadBanks();
            Shown += delegate
            {
                WindowState = FormWindowState.Maximized;
            };
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18, 22, 14, 14), ColumnCount = 3, RowCount = 3 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 520));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
            Controls.Add(root);

            var title = new Label { Text = "OJT EXAM MAKER", Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("Segoe UI", 20, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(0, 4, 0, 0) };
            root.Controls.Add(title, 0, 0);
            root.SetColumnSpan(title, 2);
            statusLabel = new Label { Text = "READY", Dock = DockStyle.Right, ForeColor = Color.FromArgb(131, 255, 192), BackColor = Color.FromArgb(16, 47, 34), Font = new Font("Segoe UI", 9, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter, Width = 120, Margin = new Padding(0, 4, 0, 12) };
            root.Controls.Add(statusLabel, 2, 0);

            var filePanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, BackColor = Color.FromArgb(18, 28, 46), Padding = new Padding(8), Margin = new Padding(0, 0, 0, 10) };
            filePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            filePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            filePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            filePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 104));
            root.Controls.Add(filePanel, 0, 1);
            root.SetColumnSpan(filePanel, 3);
            filePanel.Controls.Add(MutedLabel("문제은행"), 0, 0);
            pathBox = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(8, 17, 31), ForeColor = Color.White };
            filePanel.Controls.Add(pathBox, 1, 0);
            filePanel.Controls.Add(DarkButton("파일 선택", BrowseWorkbook), 2, 0);
            filePanel.Controls.Add(DarkButton("새로고침", delegate { LoadBanks(); }), 3, 0);

            root.Controls.Add(BuildBankPanel(), 0, 2);
            root.Controls.Add(BuildSettingPanel(), 1, 2);
            root.Controls.Add(BuildPreviewPanel(), 2, 2);
        }

        private Control BuildBankPanel()
        {
            var panel = PanelBox("공정 선택");
            bankGrid = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.FromArgb(18, 28, 46), BorderStyle = BorderStyle.None, AllowUserToAddRows = false, AllowUserToDeleteRows = false, ReadOnly = true, RowHeadersVisible = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false };
            bankGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 45, 72);
            bankGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            bankGrid.DefaultCellStyle.BackColor = Color.FromArgb(18, 28, 46);
            bankGrid.DefaultCellStyle.ForeColor = Color.White;
            bankGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(20, 81, 63);
            bankGrid.DefaultCellStyle.SelectionForeColor = Color.White;
            bankGrid.Columns.Add("name", "공정명");
            bankGrid.Columns.Add("common", "공통");
            bankGrid.Columns.Add("choice", "객관식");
            bankGrid.Columns.Add("subjective", "주관식");
            bankGrid.Columns[0].FillWeight = 220;
            bankGrid.Columns[1].FillWeight = 72;
            bankGrid.Columns[2].FillWeight = 72;
            bankGrid.Columns[3].FillWeight = 72;
            bankGrid.CellFormatting += delegate(object sender, DataGridViewCellFormattingEventArgs e)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 0 && e.Value != null)
                    bankGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = e.Value.ToString();
            };
            bankGrid.SelectionChanged += delegate { RefreshSelected(); };
            panel.Controls.Add(bankGrid);
            return panel;
        }

        private Control BuildSettingPanel()
        {
            var panel = PanelBox("시험 조건");
            var body = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(14), AutoScroll = true };
            panel.Controls.Add(body);
            body.Controls.Add(TypeRow("공통", out commonCount, out commonScore, 2, 2.5m));
            body.Controls.Add(TypeRow("객관식", out choiceCount, out choiceScore, 20, 4m));
            body.Controls.Add(TypeRow("주관식", out subjectiveCount, out subjectiveScore, 3, 5m));
            targetScore = Number(0, 1000, 100, 0);
            body.Controls.Add(LabeledNumber("목표 점수", targetScore));
            totalLabel = new Label { Text = "TOTAL 100", ForeColor = Color.FromArgb(32, 208, 132), Font = new Font("Segoe UI", 27, FontStyle.Bold), Width = 470, Height = 58, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.FromArgb(16, 36, 58), Margin = new Padding(0, 12, 0, 0) };
            body.Controls.Add(totalLabel);
            selectedLabel = new Label { Text = "공정을 선택하세요.", ForeColor = Color.White, Width = 470, Height = 64, Padding = new Padding(10), BackColor = Color.FromArgb(10, 19, 34), Margin = new Padding(0, 8, 0, 0) };
            body.Controls.Add(selectedLabel);
            answerCheck = new CheckBox { Text = "미리보기에 답안 표시", ForeColor = Color.White, Width = 470, Margin = new Padding(0, 10, 0, 0) };
            answerCheck.CheckedChanged += delegate { if (currentDoc != null) { currentDoc.ShowAnswers = answerCheck.Checked; preview.Invalidate(); } };
            body.Controls.Add(answerCheck);

            var nav = new TableLayoutPanel { Width = 470, Height = 36, ColumnCount = 3, Margin = new Padding(0, 8, 0, 0) };
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            nav.Controls.Add(DarkButton("이전", delegate { preview.MovePage(-1); UpdatePageLabel(); }), 0, 0);
            pageLabel = new Label { Text = "1 / 1", Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            nav.Controls.Add(pageLabel, 1, 0);
            nav.Controls.Add(DarkButton("다음", delegate { preview.MovePage(1); UpdatePageLabel(); }), 2, 0);
            body.Controls.Add(nav);

            body.Controls.Add(ActionButton("랜덤 미리보기", GeneratePreview, Color.FromArgb(32, 208, 132)));
            body.Controls.Add(ActionButton("문제 프린트", delegate { PrintCurrent(false, false); }, Color.FromArgb(30, 45, 72)));
            body.Controls.Add(ActionButton("답안 프린트", delegate { PrintCurrent(true, true); }, Color.FromArgb(30, 45, 72)));
            body.Controls.Add(ActionButton("문제 + 답안 프린트", delegate { PrintProblemAndAnswer(); }, Color.FromArgb(30, 45, 72)));
            return panel;
        }

        private Control BuildPreviewPanel()
        {
            var panel = PanelBox("시험지 미리보기");
            preview = new PreviewCanvas { Dock = DockStyle.Fill };
            panel.Controls.Add(preview);
            return panel;
        }

        private Panel PanelBox(string title)
        {
            var panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 28, 46), Padding = new Padding(14), Margin = new Padding(0, 0, 12, 0) };
            var label = new Label { Text = title, Dock = DockStyle.Top, Height = 34, ForeColor = Color.White, Font = new Font("맑은 고딕", 14, FontStyle.Bold) };
            panel.Controls.Add(label);
            label.BringToFront();
            return panel;
        }

        private static Label MutedLabel(string text)
        {
            return new Label { Text = text, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), Font = new Font("맑은 고딕", 9, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        }

        private Button DarkButton(string text, EventHandler click)
        {
            var b = new Button { Text = text, Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(30, 45, 72), ForeColor = Color.White, Font = new Font("맑은 고딕", 9, FontStyle.Bold), Margin = new Padding(4, 0, 0, 0) };
            b.FlatAppearance.BorderColor = Color.FromArgb(220, 230, 245);
            b.Click += click;
            return b;
        }

        private Button ActionButton(string text, EventHandler click, Color color)
        {
            var b = new Button { Text = text, Width = 470, Height = 44, FlatStyle = FlatStyle.Flat, BackColor = color, ForeColor = color == Color.FromArgb(32, 208, 132) ? Color.Black : Color.White, Font = new Font("맑은 고딕", 10, FontStyle.Bold), Margin = new Padding(0, 10, 0, 0) };
            b.FlatAppearance.BorderColor = Color.White;
            b.Click += click;
            return b;
        }

        private Control TypeRow(string label, out NumericUpDown count, out NumericUpDown score, int countValue, decimal scoreValue)
        {
            var row = new TableLayoutPanel { Width = 470, Height = 44, ColumnCount = 3 };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f));
            row.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("맑은 고딕", 12, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            count = Number(0, 500, countValue, 0);
            score = Number(0, 100, scoreValue, 1);
            row.Controls.Add(count, 1, 0);
            row.Controls.Add(score, 2, 0);
            return row;
        }

        private Control LabeledNumber(string label, NumericUpDown num)
        {
            var row = new TableLayoutPanel { Width = 470, Height = 44, ColumnCount = 2 };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            row.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, ForeColor = Color.White, Font = new Font("맑은 고딕", 11, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            row.Controls.Add(num, 1, 0);
            return row;
        }

        private NumericUpDown Number(decimal min, decimal max, decimal value, int decimals)
        {
            var n = new NumericUpDown { Minimum = min, Maximum = max, Value = value, DecimalPlaces = decimals, Increment = decimals == 0 ? 1 : .5m, TextAlign = HorizontalAlignment.Center, Width = 100, BackColor = Color.FromArgb(8, 17, 31), ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            n.ValueChanged += delegate { UpdateTotal(); };
            return n;
        }

        private void BrowseWorkbook(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel files (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|All files (*.*)|*.*";
                dlg.Title = "문제은행 엑셀 선택";
                if (File.Exists(pathBox.Text))
                    dlg.InitialDirectory = Path.GetDirectoryName(pathBox.Text);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    pathBox.Text = dlg.FileName;
                    SaveLastWorkbookPath(dlg.FileName);
                    LoadBanks();
                }
            }
        }

        private void LoadBanks()
        {
            try
            {
                SetStatus("LOADING", false);
                if (!File.Exists(pathBox.Text))
                    throw new FileNotFoundException("문제은행 파일을 찾을 수 없습니다.", pathBox.Text);
                SaveLastWorkbookPath(pathBox.Text);
                banks.Clear();
                banks.AddRange(reader.Load(pathBox.Text));
                bankGrid.Rows.Clear();
                foreach (var bank in banks)
                    bankGrid.Rows.Add(bank.Name, bank.Common, bank.Choice, bank.Subjective);
                if (bankGrid.Rows.Count > 0)
                    bankGrid.Rows[0].Selected = true;
                RefreshSelected();
                SetStatus("READY", false);
            }
            catch (Exception ex)
            {
                SetStatus("ERROR", true);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private QuestionBank SelectedBank()
        {
            if (bankGrid.CurrentRow == null || bankGrid.CurrentRow.Index < 0 || bankGrid.CurrentRow.Index >= banks.Count)
                return null;
            return banks[bankGrid.CurrentRow.Index];
        }

        private void RefreshSelected()
        {
            var bank = SelectedBank();
            selectedLabel.Text = bank == null ? "공정을 선택하세요." : bank.Name + Environment.NewLine + string.Format("공통 {0} / 객관식 {1} / 주관식 {2}", bank.Common, bank.Choice, bank.Subjective);
        }

        private void UpdateTotal()
        {
            decimal total = commonCount.Value * commonScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value;
            if (totalLabel != null)
            {
                totalLabel.Text = "TOTAL " + total.ToString("0.##");
                totalLabel.ForeColor = total == targetScore.Value ? Color.FromArgb(32, 208, 132) : Color.FromArgb(246, 180, 75);
            }
        }

        private void GeneratePreview(object sender, EventArgs e)
        {
            try
            {
                var bank = SelectedBank();
                if (bank == null)
                    throw new InvalidOperationException("공정을 선택하세요.");
                decimal total = commonCount.Value * commonScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value;
                if (total != targetScore.Value)
                    throw new InvalidOperationException(string.Format("총점이 목표 점수와 다릅니다. 현재 {0:0.##}점 / 목표 {1:0.##}점", total, targetScore.Value));
                var doc = new ExamDocument
                {
                    ProcessName = bank.Name,
                    Department = bank.Department,
                    Evaluator = bank.Evaluator,
                    JobName = bank.JobName,
                    Revision = bank.Revision,
                    IssueDate = bank.IssueDate,
                    ProductType = bank.ProductType,
                    CreatedAt = DateTime.Now,
                    ShowAnswers = answerCheck.Checked
                };
                AddRandom(doc, bank, "공통", (int)commonCount.Value, commonScore.Value);
                AddRandom(doc, bank, "객관식", (int)choiceCount.Value, choiceScore.Value);
                AddRandom(doc, bank, "주관식", (int)subjectiveCount.Value, subjectiveScore.Value);
                doc.Questions.Sort((a, b) => random.Next(-1, 2));
                currentDoc = doc;
                preview.Document = doc;
                preview.PageIndex = 0;
                preview.Invalidate();
                UpdatePageLabel();
                SetStatus("READY", false);
            }
            catch (Exception ex)
            {
                SetStatus("ERROR", true);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddRandom(ExamDocument doc, QuestionBank bank, string type, int count, decimal score)
        {
            var candidates = bank.Questions.Where(q => q.ExamType == type).OrderBy(q => random.Next()).Take(count).ToList();
            if (candidates.Count < count)
                throw new InvalidOperationException(string.Format("{0} 문제가 부족합니다. 요청 {1}문항 / 보유 {2}문항", type, count, candidates.Count));
            foreach (var source in candidates)
            {
                var q = new Question { SourceSheet = source.SourceSheet, SourceNo = source.SourceNo, ExamType = source.ExamType, Text = source.Text, Answer = source.Answer, Score = score };
                q.Images.AddRange(source.Images);
                doc.Questions.Add(q);
            }
        }

        private static void ApplyHeaderInfo(ExamDocument doc, string processName)
        {
            doc.JobName = processName;
            doc.ProductType = "";
            if (processName.EndsWith(" 일반용", StringComparison.Ordinal))
            {
                doc.JobName = processName.Substring(0, processName.Length - " 일반용".Length);
                doc.ProductType = "일반용";
            }
            else if (processName.EndsWith(" 전장용", StringComparison.Ordinal))
            {
                doc.JobName = processName.Substring(0, processName.Length - " 전장용".Length);
                doc.ProductType = "전장용";
            }
        }

        private void UpdatePageLabel()
        {
            pageLabel.Text = string.Format("{0} / {1}", preview.PageIndex + 1, preview.PageCount);
        }

        private void PrintCurrent(bool showAnswers, bool answerOnly)
        {
            try
            {
                if (currentDoc == null)
                    GeneratePreview(null, EventArgs.Empty);
                if (currentDoc == null)
                    return;

                using (var printDoc = new PrintDocument())
                {
                    printDoc.DocumentName = "OJT 시험지";
                    printDoc.DefaultPageSettings.Landscape = true;
                    printDoc.DefaultPageSettings.Margins = new Margins(20, 20, 20, 20);
                    int page = 0;
                    int totalPages = answerOnly ? 1 : preview.PageCount;
                    printDoc.PrintPage += delegate(object sender, PrintPageEventArgs e)
                    {
                        if (answerOnly)
                            DrawAnswerSheet(e.Graphics, e.MarginBounds);
                        else
                            preview.DrawForPrint(e.Graphics, e.MarginBounds, page, showAnswers);
                        page++;
                        e.HasMorePages = page < totalPages;
                    };
                    using (var dlg = new PrintDialog { Document = printDoc, UseEXDialog = true })
                    {
                        if (dlg.ShowDialog(this) == DialogResult.OK)
                            printDoc.Print();
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus("ERROR", true);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PrintProblemAndAnswer()
        {
            try
            {
                if (currentDoc == null)
                    GeneratePreview(null, EventArgs.Empty);
                if (currentDoc == null)
                    return;

                using (var printDoc = new PrintDocument())
                {
                    printDoc.DocumentName = "OJT 시험지 + 답안지";
                    printDoc.DefaultPageSettings.Landscape = true;
                    printDoc.DefaultPageSettings.Margins = new Margins(20, 20, 20, 20);
                    int page = 0;
                    int problemPages = preview.PageCount;
                    printDoc.PrintPage += delegate(object sender, PrintPageEventArgs e)
                    {
                        if (page < problemPages)
                            preview.DrawForPrint(e.Graphics, e.MarginBounds, page, false);
                        else
                            DrawAnswerSheet(e.Graphics, e.MarginBounds);
                        page++;
                        e.HasMorePages = page < problemPages + 1;
                    };
                    using (var dlg = new PrintDialog { Document = printDoc, UseEXDialog = true })
                    {
                        if (dlg.ShowDialog(this) == DialogResult.OK)
                            printDoc.Print();
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus("ERROR", true);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DrawAnswerSheet(Graphics g, Rectangle bounds)
        {
            g.FillRectangle(Brushes.White, bounds);
            using (var title = new Font("맑은 고딕", 18, FontStyle.Bold))
            using (var font = new Font("맑은 고딕", 10))
            using (var bold = new Font("맑은 고딕", 10, FontStyle.Bold))
            using (var pen = new Pen(Color.Black, 1))
            {
                g.DrawString("OJT 답안지", title, Brushes.Black, bounds.Left, bounds.Top);
                g.DrawString(currentDoc.ProcessName + " / " + currentDoc.CreatedAt.ToString("yyyy-MM-dd"), font, Brushes.Black, bounds.Left, bounds.Top + 34);
                int cols = 3;
                int colW = bounds.Width / cols;
                int rowH = 26;
                int y0 = bounds.Top + 70;
                for (int i = 0; i < currentDoc.Questions.Count; i++)
                {
                    int col = i / 24;
                    int row = i % 24;
                    if (col >= cols)
                        break;
                    int x = bounds.Left + col * colW;
                    int y = y0 + row * rowH;
                    var rect = new Rectangle(x, y, colW - 12, rowH);
                    g.DrawRectangle(pen, rect);
                    g.DrawString((i + 1).ToString(), bold, Brushes.Black, rect.X + 6, rect.Y + 4);
                    g.DrawString(PreviewCanvas.ApplyAnswer(currentDoc.Questions[i]), font, Brushes.Black, rect.X + 54, rect.Y + 4);
                }
            }
        }

        private void SetStatus(string text, bool error)
        {
            statusLabel.Text = text;
            statusLabel.BackColor = error ? Color.FromArgb(59, 17, 17) : Color.FromArgb(16, 47, 34);
            statusLabel.ForeColor = error ? Color.FromArgb(255, 176, 176) : Color.FromArgb(131, 255, 192);
        }

        private static string ConfigPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OJT_Exam_Maker");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "last_workbook.txt");
        }

        private static void SaveLastWorkbookPath(string path)
        {
            try
            {
                File.WriteAllText(ConfigPath(), path, Encoding.UTF8);
            }
            catch
            {
            }
        }

        private static string DefaultWorkbookPath2()
        {
            try
            {
                string cfg = ConfigPath();
                if (File.Exists(cfg))
                {
                    string saved = File.ReadAllText(cfg, Encoding.UTF8).Trim();
                    if (File.Exists(saved))
                        return saved;
                }
            }
            catch
            {
            }

            var dirs = new List<string>();
            dirs.Add(Environment.CurrentDirectory);
            dirs.Add(AppDomain.CurrentDomain.BaseDirectory);
            var parent = Directory.GetParent(Environment.CurrentDirectory);
            if (parent != null)
                dirs.Add(parent.FullName);
            dirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            foreach (string dir in dirs.Where(Directory.Exists).Distinct())
            {
                string file = Directory.GetFiles(dir, "*.xlsm").FirstOrDefault(f => !Path.GetFileName(f).StartsWith("~$"));
                if (file != null)
                    return file;
            }
            return Path.Combine(Environment.CurrentDirectory, "OJT 시험 문제.xlsm");
        }

        private static string DefaultWorkbookPath()
        {
            var candidates = new List<string>();
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "OJT 시험 문제.xlsm"));
            candidates.Add(Path.Combine(Environment.CurrentDirectory, "OJT 시험 문제.xlsm"));
            var parent = Directory.GetParent(Environment.CurrentDirectory);
            if (parent != null)
                candidates.Add(Path.Combine(parent.FullName, "OJT 시험 문제.xlsm"));
            foreach (var path in candidates)
                if (File.Exists(path))
                    return path;
            return candidates[0];
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--self-test")
            {
                string path = args.Length > 1 ? args[1] : Path.Combine(Environment.CurrentDirectory, "OJT 시험 문제.xlsm");
                var banks = new XlsmQuestionReader().Load(path);
                int imageCount = banks.Sum(b => b.Questions.Sum(q => q.Images.Count));
                Console.WriteLine("OK banks={0} questions={1} images={2}", banks.Count, banks.Sum(b => b.Questions.Count), imageCount);
                foreach (var bank in banks.Take(5))
                    Console.WriteLine("{0}: 공통 {1}, 객관식 {2}, 주관식 {3}", bank.Name, bank.Common, bank.Choice, bank.Subjective);
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
