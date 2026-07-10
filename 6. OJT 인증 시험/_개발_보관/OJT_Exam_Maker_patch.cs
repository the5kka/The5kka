using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
        public int Common { get { return Questions.Count(q => q.Type == "\uACF5\uD1B5"); } }
        public int Choice { get { return Questions.Count(q => q.Type == "\uAC1D\uAD00\uC2DD"); } }
        public int Subjective { get { return Questions.Count(q => q.Type == "\uC8FC\uAD00\uC2DD"); } }
    }

    internal sealed class ExamDoc
    {
        public Bank Bank;
        public readonly List<Question> Questions = new List<Question>();
        public bool ShowAnswers;
        public string UserName = "";
        public string EvalDate = "";
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
                    var images = ImagesByRow(zip, sheet.Path);
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
                        if (type != "\uACF5\uD1B5" && type != "\uAC1D\uAD00\uC2DD" && type != "\uC8FC\uAD00\uC2DD")
                            continue;
                        double score;
                        if (!double.TryParse(Clean(Get(cells, row, header.ScoreCol)), out score))
                            score = type == "\uACF5\uD1B5" ? 2.5 : type == "\uAC1D\uAD00\uC2DD" ? 4.0 : 5.0;
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
                        else if (LooksLikeImageQuestion(questionText))
                            AddNearbyImages(q, images, cells, header, row);
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

        static Dictionary<int, List<Image>> ImagesByRow(ZipArchive zip, string sheetPath)
        {
            var result = new Dictionary<int, List<Image>>();
            var sheetRels = zip.GetEntry(PathToRels(sheetPath));
            if (sheetRels == null) return result;
            string drawingPath = null;
            using (var s = sheetRels.Open())
            {
                var doc = XDocument.Load(s);
                var rel = doc.Root.Elements(PkgRelNs + "Relationship").FirstOrDefault(r => ((string)r.Attribute("Type") ?? "").EndsWith("/drawing"));
                if (rel != null)
                    drawingPath = ResolvePart(sheetPath, (string)rel.Attribute("Target"));
            }
            if (drawingPath == null || zip.GetEntry(drawingPath) == null) return result;
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
                    var group = anchor.Element(DrawNs + "grpSp");
                    if (blip == null && group != null)
                    {
                        Image shapeImage = RenderGroupShape(group);
                        if (shapeImage != null)
                        {
                            if (!result.ContainsKey(row)) result[row] = new List<Image>();
                            result[row].Add(shapeImage);
                        }
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
                        if (!result.ContainsKey(row)) result[row] = new List<Image>();
                        result[row].Add(new Bitmap(img));
                    }
                }
            }
            return result;
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
        static void AddNearbyImages(Question q, Dictionary<int, List<Image>> images, Dictionary<Tuple<int, int>, string> cells, Header header, int row)
        {
            int[] offsets = { 1, -1, 2, -2 };
            foreach (int offset in offsets)
            {
                int target = row + offset;
                if (target <= 0 || HasQuestionNumberBetween(cells, header, row, target))
                    continue;
                List<Image> found;
                if (images.TryGetValue(target, out found))
                {
                    q.Images.AddRange(found);
                    return;
                }
            }
        }
        static bool HasQuestionNumberBetween(Dictionary<Tuple<int, int>, string> cells, Header header, int row, int target)
        {
            int start = Math.Min(row, target) + 1;
            int end = Math.Max(row, target) - 1;
            for (int r = start; r <= end; r++)
                if (Clean(Get(cells, r, header.NoCol)).Length > 0)
                    return true;
            return false;
        }
        static bool LooksLikeImageQuestion(string text)
        {
            text = text ?? "";
            return text.IndexOf("\u2460", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("\u2461", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("\u2462", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("\u2463", StringComparison.Ordinal) >= 0;
        }
        static string Norm(string s) { return Regex.Replace(Clean(s), @"\s+", "").ToLowerInvariant(); }
        static string Get(Dictionary<Tuple<int, int>, string> cells, int row, int col)
        {
            if (row <= 0 || col <= 0) return "";
            string value; return cells.TryGetValue(Tuple.Create(row, col), out value) ? value : "";
        }
        static int ToInt(string s) { int n; return int.TryParse(s, out n) ? n : 0; }
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
            e.Graphics.FillRectangle(Brushes.White, page);
            e.Graphics.DrawRectangle(Pens.LightGray, page);
            Rectangle content = PrintContentRect(page);
            if (PageIndex >= QuestionPageCount)
                DrawAnswerPage(e.Graphics, content);
            else
                DrawPage(e.Graphics, content, PageIndex, Doc.ShowAnswers);
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
                if (i > 0 && used + h > limit)
                {
                    starts.Add(i);
                    used = 0;
                    limit = 760;
                }
                used += h;
            }
            int maxPage = Math.Max(1, starts.Count) + 1;
            if (PageIndex >= maxPage) PageIndex = maxPage - 1;
        }
        int EndIndex(int page) { return page + 1 < starts.Count ? starts[page + 1] : Doc.Questions.Count; }
        public static string AnswerMark(Question q)
        {
            if (q.Type == "\uAC1D\uAD00\uC2DD" || q.Type == "\uACF5\uD1B5")
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
            string display = FormatQuestionText(q);
            if (q.Images.Count > 0)
                display = StripImageChoiceLines(display);
            int lines = Math.Max(1, (display ?? "").Count(c => c == '\n') + 1);
            int textLines = Math.Max(lines, (display ?? "").Length / 92 + 1);
            int image = q.Images.Count == 0 ? 0 : 132;
            bool subjType = q.Type == "\uC8FC\uAD00\uC2DD";
            if (subjType)
                return Math.Max(58, textLines * 17 + image + 8);
            if (q.Images.Count > 0)
                return Math.Max(156, textLines * 13 + image + 10);
            return Math.Max(34, textLines * 13 + 4);
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
                    y = Y(228);
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
            Cell(g, "\uC9C1 \uBB34 \uBA85", bold, new Rectangle(X(16), Y(126 + dy), W(230), H(26)), true);
            Cell(g, b.JobName, font, new Rectangle(X(246), Y(126 + dy), W(470), H(26)), true);
            Cell(g, "\uC131 \uBA85", bold, new Rectangle(X(716), Y(126 + dy), W(90), H(26)), true);
            Cell(g, Doc.UserName, font, new Rectangle(X(806), Y(126 + dy), W(300), H(26)), true);
            Cell(g, "\uD3C9\uAC00 \uBC29\uBC95", bold, new Rectangle(X(16), Y(152 + dy), W(230), H(26)), true);
            Cell(g, "\uC2DC\uD5D8 \uD3C9\uAC00, \uACB0\uACFC \uBCF4\uACE0\uC11C, \uC9C1\uBB34 \uD3C9\uAC00, \uAE30\uD0C0 \uBC29\uBC95 (          )", font, new Rectangle(X(246), Y(152 + dy), W(860), H(26)), true);
            Cell(g, "\uAC1C\uC815 \uCC28\uC218", bold, new Rectangle(X(16), Y(178 + dy), W(230), H(26)), true);
            Cell(g, b.Revision, font, new Rectangle(X(246), Y(178 + dy), W(160), H(26)), true);
            Cell(g, "\uC81C\uC815\uC77C", bold, new Rectangle(X(406), Y(178 + dy), W(160), H(26)), true);
            Cell(g, b.IssueDate, font, new Rectangle(X(566), Y(178 + dy), W(170), H(26)), true);
            Cell(g, "\uC720 \uD615", bold, new Rectangle(X(736), Y(178 + dy), W(150), H(26)), true);
            Cell(g, b.ProductType, font, new Rectangle(X(886), Y(178 + dy), W(220), H(26)), true);
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

            h(126 + dy, 16, 1106); h(152 + dy, 16, 1106); h(178 + dy, 16, 1106); h(204 + dy, 16, 1106);
            v(16, 126 + dy, 204 + dy); v(1106, 126 + dy, 204 + dy);
            v(246, 126 + dy, 204 + dy); v(716, 126 + dy, 152 + dy); v(806, 126 + dy, 152 + dy);
            v(406, 178 + dy, 204 + dy); v(566, 178 + dy, 204 + dy); v(736, 178 + dy, 204 + dy); v(886, 178 + dy, 204 + dy);
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
                string imageText = StripImageChoiceLines(text);
                string choiceText = ImageChoiceLines(text);
                int imgH = 98;
                int gap = 42;
                int imgY = rect.Y + 36;
                int imgW = 96;
                int startX = rect.X + 42;
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
                    string[] marks = { "\u2460", "\u2461", "\u2462", "\u2463" };
                    g.DrawString(marks[idx], font, brush, x - 20, imgY + imgH / 2 - 7);
                    g.DrawImage(img, x, imgY, iw, ih);
                }
                textRect.Height = Math.Max(22, imgY - textRect.Y - 4);
                DrawSpacedText(g, imageText, font, brush, textRect, 0);
                if (choiceText.Trim().Length > 0)
                {
                    var choiceRect = new Rectangle(rect.X + 42, imgY + imgH + 4, rect.Width - 80, Math.Max(16, rect.Bottom - imgY - imgH - 6));
                    DrawSpacedText(g, choiceText, font, brush, choiceRect, 0);
                }
                if (answers)
                {
                    using (var red = new SolidBrush(Color.FromArgb(200, 0, 0)))
                        g.DrawString(AnswerMark(q), bold, red, rect.Right - 70, rect.Y + 5);
                }
                return;
            }
            DrawSpacedText(g, text, font, brush, textRect, q.Type == "\uC8FC\uAD00\uC2DD" ? 2 : 0);
            if (answers)
            {
                using (var red = new SolidBrush(Color.FromArgb(200, 0, 0)))
                    g.DrawString(AnswerMark(q), bold, red, rect.Right - 70, rect.Y + 5);
            }
        }
        static string StripImageChoiceLines(string text)
        {
            var kept = new List<string>();
            foreach (string line in (text ?? "").Split('\n'))
            {
                if (Regex.IsMatch(line.Trim(), @"^[①②③④⑤⑥]\s+.+$"))
                    continue;
                kept.Add(line);
            }
            return string.Join("\n", kept);
        }
        static string ImageChoiceLines(string text)
        {
            var kept = new List<string>();
            foreach (string line in (text ?? "").Split('\n'))
                if (Regex.IsMatch(line.Trim(), @"^[①②③④⑤⑥]\s+.+$"))
                    kept.Add(line);
            return string.Join("\n", kept);
        }
        static void DrawSpacedText(Graphics g, string text, Font font, Brush brush, Rectangle rect, int extra)
        {
            int y = rect.Y;
            foreach (string line in (text ?? "").Split('\n'))
            {
                if (line.Trim().Length == 0)
                {
                    y += Math.Max(2, extra + 2);
                    continue;
                }
                g.DrawString(line, font, brush, rect.X, y);
                y += font.Height + extra;
                if (y > rect.Bottom) break;
            }
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
            string score = " (" + q.Score.ToString("0.##") + "\uC810)";
            string text = q.Text ?? "";
            int lineBreak = text.IndexOf('\n');
            string first = lineBreak >= 0 ? text.Substring(0, lineBreak) : text;
            string rest = lineBreak >= 0 ? text.Substring(lineBreak) : "";
            if (first.Contains("(" + q.Score.ToString("0.##") + "\uC810)"))
                return text;
            int qmark = first.IndexOf('?');
            if (qmark >= 0)
                first = first.Insert(qmark + 1, score);
            else
                first = first + score;
            return first + rest;
        }
    }

    internal sealed class MainForm : Form
    {
        readonly Reader reader = new Reader();
        readonly Random random = new Random();
        readonly List<Bank> banks = new List<Bank>();
        TextBox pathBox;
        TextBox userNameBox, evalDateBox;
        DataGridView grid;
        NumericUpDown commonCount, choiceCount, subjectiveCount, commonScore, choiceScore, subjectiveScore, targetScore;
        Label totalLabel, selectedLabel, pageLabel, status;
        CheckBox answerCheck;
        Preview preview;
        ExamDoc doc;

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
        void BuildUi()
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
            grid.Columns.Add("name", "\uACF5\uC815\uBA85"); grid.Columns.Add("common", "\uACF5\uD1B5"); grid.Columns.Add("choice", "\uAC1D\uAD00\uC2DD"); grid.Columns.Add("subjective", "\uC8FC\uAD00\uC2DD");
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
            var conditionHeader = new TableLayoutPanel { Width = 470, Height = 24, ColumnCount = 3, Margin = new Padding(0, 0, 0, 2) };
            conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f)); conditionHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5f));
            conditionHeader.Controls.Add(new Label { Text = "", Dock = DockStyle.Fill }, 0, 0);
            conditionHeader.Controls.Add(new Label { Text = "\uBB38\uC81C\uC218", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter }, 1, 0);
            conditionHeader.Controls.Add(new Label { Text = "\uC810\uC218", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(176, 190, 215), Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter }, 2, 0);
            body.Controls.Add(conditionHeader);
            body.Controls.Add(Row("\uACF5\uD1B5", out commonCount, out commonScore, 2, 2.5m, true));
            body.Controls.Add(Row("\uAC1D\uAD00\uC2DD", out choiceCount, out choiceScore, 20, 4m, true));
            body.Controls.Add(Row("\uC8FC\uAD00\uC2DD", out subjectiveCount, out subjectiveScore, 3, 5m, true));
            targetScore = Num(0, 1000, 100, 0); body.Controls.Add(TargetRow());
            totalLabel = new Label { Text = "TOTAL 100", Width = 470, Height = 52, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(32, 208, 132), BackColor = Color.FromArgb(16, 36, 58), Font = new Font("Segoe UI", 25, FontStyle.Bold), Margin = new Padding(0, 8, 0, 0) };
            body.Controls.Add(totalLabel);
            body.Controls.Add(InputRow("\uC131\uBA85", out userNameBox, ""));
            body.Controls.Add(InputRow("\uD3C9\uAC00 \uC77C\uC2DC", out evalDateBox, DateTime.Today.ToString("yyyy.MM.dd")));
            selectedLabel = new Label { Text = "\uACF5\uC815\uC744 \uC120\uD0DD\uD558\uC138\uC694.", Width = 470, Height = 64, ForeColor = Color.White, BackColor = Color.FromArgb(10, 19, 34), Padding = new Padding(10), Margin = new Padding(0, 8, 0, 0) };
            body.Controls.Add(selectedLabel);
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
        NumericUpDown Num(decimal min, decimal max, decimal value, int dec) { var n = new NumericUpDown { Minimum = min, Maximum = max, Value = value, DecimalPlaces = dec, Increment = dec == 0 ? 1 : .5m, TextAlign = HorizontalAlignment.Center, Dock = DockStyle.Fill, BackColor = Color.FromArgb(8, 17, 31), ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold) }; n.ValueChanged += delegate { UpdateTotal(); }; return n; }
        void Browse(object s, EventArgs e) { using (var d = new OpenFileDialog { Filter = "Excel files (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|All files (*.*)|*.*", Title = "\uBB38\uC81C\uC740\uD589 \uC5D1\uC140 \uC120\uD0DD" }) { if (File.Exists(pathBox.Text)) d.InitialDirectory = Path.GetDirectoryName(pathBox.Text); if (d.ShowDialog(this) == DialogResult.OK) { pathBox.Text = d.FileName; SaveLastWorkbook(d.FileName); LoadBanks(); } } }
        void LoadBanks()
        {
            try
            {
                Status("LOADING", false); doc = null; if (preview != null) { preview.Doc = null; preview.PageIndex = 0; preview.Invalidate(); } banks.Clear(); banks.AddRange(reader.Load(pathBox.Text)); SaveLastWorkbook(pathBox.Text); grid.Rows.Clear();
                foreach (Bank b in banks) grid.Rows.Add(b.Name, b.Common, b.Choice, b.Subjective);
                if (grid.Rows.Count > 0) grid.Rows[0].Selected = true;
                RefreshSelected(); UpdateTotal(); UpdatePage(); Status("READY", false);
            }
            catch (Exception ex) { Status("ERROR", true); MessageBox.Show(this, ex.Message, "\uC624\uB958", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        Bank SelectedBank() { return grid.CurrentRow == null || grid.CurrentRow.Index < 0 || grid.CurrentRow.Index >= banks.Count ? null : banks[grid.CurrentRow.Index]; }
        void RefreshSelected() { Bank b = SelectedBank(); selectedLabel.Text = b == null ? "\uACF5\uC815\uC744 \uC120\uD0DD\uD558\uC138\uC694." : b.Name + Environment.NewLine + string.Format("\uACF5\uD1B5 {0} / \uAC1D\uAD00\uC2DD {1} / \uC8FC\uAD00\uC2DD {2}", b.Common, b.Choice, b.Subjective); }
        void UpdateTotal() { if (totalLabel == null) return; decimal total = commonCount.Value * commonScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value; totalLabel.Text = "TOTAL " + total.ToString("0.##"); totalLabel.ForeColor = total == targetScore.Value ? Color.FromArgb(32, 208, 132) : Color.FromArgb(246, 180, 75); }
        void Generate(object s, EventArgs e)
        {
            try
            {
                Bank b = SelectedBank(); if (b == null) throw new InvalidOperationException("\uACF5\uC815\uC744 \uC120\uD0DD\uD558\uC138\uC694.");
                decimal total = commonCount.Value * commonScore.Value + choiceCount.Value * choiceScore.Value + subjectiveCount.Value * subjectiveScore.Value;
                if (total != targetScore.Value) throw new InvalidOperationException(string.Format("\uCD1D\uC810\uC774 \uBAA9\uD45C \uC810\uC218\uC640 \uB2E4\uB985\uB2C8\uB2E4. \uD604\uC7AC {0:0.##}\uC810 / \uBAA9\uD45C {1:0.##}\uC810", total, targetScore.Value));
                doc = new ExamDoc { Bank = b, ShowAnswers = answerCheck.Checked, UserName = userNameBox.Text.Trim(), EvalDate = evalDateBox.Text.Trim() };
                Pick(doc, b, "\uACF5\uD1B5", (int)commonCount.Value, (double)commonScore.Value);
                Pick(doc, b, "\uAC1D\uAD00\uC2DD", (int)choiceCount.Value, (double)choiceScore.Value);
                Pick(doc, b, "\uC8FC\uAD00\uC2DD", (int)subjectiveCount.Value, (double)subjectiveScore.Value);
                preview.Doc = doc; preview.PageIndex = 0; preview.Invalidate(); UpdatePage(); Status("READY", false);
            }
            catch (Exception ex) { Status("ERROR", true); MessageBox.Show(this, ex.Message, "\uC624\uB958", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        void Pick(ExamDoc d, Bank b, string type, int count, double score)
        {
            var cand = b.Questions.Where(q => q.Type == type).OrderBy(q => random.Next()).Take(count).ToList();
            if (cand.Count < count) throw new InvalidOperationException(string.Format("{0} \uBB38\uC81C\uAC00 \uBD80\uC871\uD569\uB2C8\uB2E4. \uC694\uCCAD {1}\uBB38\uD56D / \uBCF4\uC720 {2}\uBB38\uD56D", type, count, cand.Count));
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
                    dlg.ShowDialog(this);
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
                DrawInfoCell(g, gridPen, head, x + labelW + valueW, y, labelW, infoH, "\uC2DC\uD5D8\uC77C\uC790");
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
        void Status(string t, bool err) { status.Text = t; status.BackColor = err ? Color.FromArgb(59, 17, 17) : Color.FromArgb(16, 47, 34); status.ForeColor = err ? Color.FromArgb(255, 176, 176) : Color.FromArgb(131, 255, 192); }
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
            if (args.Length > 0 && args[0] == "--self-test")
            {
                string f = args.Length > 1 ? args[1] : MainFormDefaultWorkbook();
                var banks = new Reader().Load(f);
                Console.WriteLine("banks={0} questions={1} images={2}", banks.Count, banks.Sum(b => b.Questions.Count), banks.Sum(b => b.Questions.Sum(q => q.Images.Count)));
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

