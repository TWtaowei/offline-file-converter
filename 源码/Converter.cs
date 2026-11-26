// 本项目自有代码及其修改版仅供非商业使用，禁止转售及企业内部业务使用。见根目录 LICENSE。
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LocalFiles {
    static class Converter {
        internal static readonly string[] Documents = { ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp", ".rtf" };
        internal static readonly string[] Pictures = { ".png", ".jpg", ".jpeg", ".bmp" };
        internal static readonly string[] Inputs = Documents.Concat(Pictures).Concat(new[] { ".pdf", ".txt", ".csv", ".tsv", ".webp" }).ToArray();
        internal static readonly Dictionary<string,string> Labels = new Dictionary<string,string> {
            {"pdf","PDF 文档 (.pdf)"}, {"docx","Word 文档 (.docx)"}, {"xlsx","Excel 工作簿 (.xlsx)"},
            {"pptx","PowerPoint 演示文稿 (.pptx)"}, {"docx-layout","Word · 按页保留版式 (.docx)"},
            {"pptx-edit","PowerPoint · 可编辑文字和表格 (.pptx)"}, {"txt","纯文本 (.txt)"},
            {"csv","逗号分隔表格 (.csv)"}, {"tsv","制表符分隔表格 (.tsv)"}, {"odt","开放文本文档 (.odt)"},
            {"ods","开放工作簿 (.ods)"}, {"odp","开放演示文稿 (.odp)"}, {"rtf","富文本 (.rtf)"},
            {"doc","Word 97–2003 (.doc)"}, {"xls","Excel 97–2003 (.xls)"}, {"ppt","PowerPoint 97–2003 (.ppt)"},
            {"png","PNG 图片 (.png)"}, {"jpg","JPEG 图片 (.jpg)"}, {"webp","WebP 图片 (.webp)"}, {"bmp","位图 (.bmp)"}
        };
        internal static string Source(string file) {
            string ext = Path.GetExtension(file).TrimStart('.').ToLowerInvariant();
            return ext == "jpeg" ? "jpg" : ext;
        }
        internal static string Extension(string format) { return format.Split('-')[0]; }
        internal static bool ImageSource(string file) { return new[] {"png","jpg","bmp","webp"}.Contains(Source(file)); }
        internal static readonly string Engine = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LibreOffice", "program", "LocalRender.exe");
        internal static bool Supports(string file, string format) {
            if (format == null || !Labels.ContainsKey(format) || !Inputs.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase)) return false;
            string source = Source(file);
            if (format == "docx-layout") return source == "pdf";
            if (format == "pptx-edit") return !ImageSource(file) && !new[] {"ppt","pptx","odp"}.Contains(source);
            if (ImageSource(file) && new[] {"txt","csv","tsv"}.Contains(format)) return false;
            return source != format || (ImageSource(file) && (format == "jpg" || format == "png"));
        }
        internal static string Describe(IEnumerable<string> files, string mode) {
            var types = files.Select(Source).ToArray();
            if (mode == null) return "先添加文件，再选择目标格式。";
            if (mode == "docx-layout") return "每页 PDF 嵌入一张整页图片，保留排版。";
            if (mode == "pptx-edit") return "按内容生成可编辑的文本框和表格，较长内容会分到多页。";
            if (new[]{"pptx","ppt","odp"}.Contains(mode) && types.Any(t=>!new[]{"pptx","ppt","odp"}.Contains(t))) return "跨类型转换时，每个打印页生成一张图片幻灯片。";
            if (mode == "csv" || mode == "tsv") return "导出单元格值，多张工作表或多页分别保存到一个文件夹。";
            if (new[]{"png","jpg","webp","bmp"}.Contains(mode)) return "逐页导出图片，多页结果集中保存在一个文件夹。";
            if (types.Contains("pdf") && new[]{"docx","xlsx","xls","ods","odt","doc","rtf"}.Contains(mode)) return "提取可编辑的文字和可识别表格，纯扫描页作为图片嵌入。";
            if (mode == "txt") return "导出可提取的文字和单元格内容。纯图片文字需要 OCR。";
            if (mode == "pdf") return "按文档的打印版式导出 PDF。";
            if (new[]{"xlsx","xls","ods"}.Contains(mode)) return "跨类型转换时，文字和表格写入单元格，图片嵌入工作表。";
            if (new[]{"docx","doc","rtf","odt"}.Contains(mode)) return "同类文档直接转存，跨类型时整理文字、表格及图片。";
            return "使用本地排版引擎转存演示文稿。";
        }
        static void RequireLocal(string path) {
            var root = Path.GetPathRoot(path);
            if (path.StartsWith(@"\\") || new DriveInfo(root).DriveType == DriveType.Network)
                throw new IOException("请选择此电脑磁盘上的文件和目录。");
        }
        internal static string Run(string input, string output, string format) {
            input = Path.GetFullPath(input);
            output = Path.GetFullPath(output);
            RequireLocal(input);
            RequireLocal(output);
            if (!Supports(input, format)) throw new IOException("此文件类型与转换方式不匹配。");
            if ((File.GetAttributes(input) & (FileAttributes)0x441000) != 0)
                throw new IOException("请先将此文件完整保存在本地磁盘。");
            Directory.CreateDirectory(output);
            string stage = Path.Combine(output, ".local-convert-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stage);
            try {
                string stem = Path.GetFileNameWithoutExtension(input);
                string suffix = Extension(format);
                string produced = Path.Combine(stage, stem + "." + suffix);
                if (format == "pdf" && Documents.Contains(Path.GetExtension(input), StringComparer.OrdinalIgnoreCase)) ExportDocument(input, produced);
                else if ((format == "png" || format == "jpg") && Pictures.Contains(Path.GetExtension(input),StringComparer.OrdinalIgnoreCase)) ExportImage(input, produced, format);
                else ExportOther(input, format, produced, stage);
                string multiple = Path.Combine(stage, "multiple");
                bool asFolder = Directory.Exists(multiple);
                if (!asFolder && !File.Exists(produced)) throw new IOException("未生成转换结果。");
                // File.Move never overwrites an existing document, even between concurrent runs.
                for (int index = 1; ; index++) {
                    string filename = stem + (asFolder ? "_" + suffix.ToUpperInvariant() : "") + (index == 1 ? "" : " (" + index + ")") + (asFolder ? "" : "." + suffix);
                    string target = Path.Combine(output, filename);
                    try {
                        if (asFolder) Directory.Move(multiple, target); else File.Move(produced, target);
                        return target;
                    }
                    catch (IOException) { if (!File.Exists(target) && !Directory.Exists(target)) throw; }
                }
            } finally {
                try { Directory.Delete(stage, true); } catch (IOException) {} catch (UnauthorizedAccessException) {}
            }
        }
        static void ExportOther(string input, string format, string produced, string stage) {
            string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "formats", "convert.py");
            string legacyPython = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Python38", "python.exe");
            string python = File.Exists(legacyPython) ? legacyPython : Path.Combine(Path.GetDirectoryName(Engine),"python.exe");
            string error = Path.Combine(stage, Guid.NewGuid().ToString("N")+".log");
            string args = "-I -B " + string.Join(" ", new[]{script,input,format,produced,error}.Select(p => "\"" + p + "\""));
            var start = new ProcessStartInfo(python,args) { UseShellExecute=false, CreateNoWindow=true, WorkingDirectory=Path.GetDirectoryName(Engine) };
            using (var child = Process.Start(start)) {
                if (!child.WaitForExit(1200000)) {
                    using (var stop = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory,"taskkill.exe"),
                        "/PID "+ child.Id +" /T /F") { UseShellExecute=false, CreateNoWindow=true })) stop.WaitForExit();
                    throw new IOException("转换超时。");
                }
                if (child.ExitCode != 0) throw new IOException(File.Exists(error) ? File.ReadAllText(error,Encoding.UTF8) : "转换组件启动失败。请完整解压工具包。");
            }
        }
        static void ExportDocument(string input, string produced) {
            if (!File.Exists(Engine)) throw new FileNotFoundException("找不到转换引擎。请完整解压工具包，并保留旁边的 LibreOffice 文件夹。");
            string profile = Path.Combine(Path.GetTempPath(), "lfc-" + Guid.NewGuid().ToString("N"));
            try { ExportWithProfile(input, produced, profile); }
            finally {
                if (Directory.Exists(profile)) {
                    try { Directory.Delete(profile, true); } catch (IOException) {} catch (UnauthorizedAccessException) {}
                }
            }
        }
        static void ExportWithProfile(string input, string produced, string profile) {
            string config = Path.Combine(profile, "user");
            Directory.CreateDirectory(config);
            File.WriteAllText(Path.Combine(config, "registrymodifications.xcu"),
                "<oor:items xmlns:oor=\"http://openoffice.org/2001/registry\">" +
                "<item oor:path=\"/org.openoffice.Office.Common/Security/Scripting\"><prop oor:name=\"MacroSecurityLevel\" oor:op=\"fuse\"><value>3</value></prop><prop oor:name=\"DisableMacrosExecution\" oor:op=\"fuse\"><value>true</value></prop></item>" +
                "<item oor:path=\"/org.openoffice.Office.Jobs/Jobs/['UpdateCheck']/Arguments\"><prop oor:name=\"AutoCheckEnabled\" oor:op=\"fuse\"><value>false</value></prop></item>" +
                "</oor:items>");
            string engineDir = Path.GetDirectoryName(Engine);
            // Writer exposes a library export. For Calc and Impress use their desktop export filters.
            string extension = Path.GetExtension(input).ToLowerInvariant();
            bool writer = new[] { ".doc", ".docx", ".rtf", ".odt" }.Contains(extension);
            string program = writer ? Engine : Path.Combine(engineDir, "soffice.com");
            string arguments = writer ? "\"" + engineDir + "\" \"" + profile + "\" \"" + input + "\" \"" + produced + "\"" :
                "\"-env:UserInstallation=" + new Uri(profile).AbsoluteUri + "\" --headless --nologo --norestore --nodefault --convert-to pdf --outdir \"" + Path.GetDirectoryName(produced) + "\" \"" + input + "\"";
            var start = new ProcessStartInfo(program, arguments) {
                UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(Engine)
            };
            using (var process = Process.Start(start)) {
                if (!process.WaitForExit(300000)) {
                    using (var stop = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "taskkill.exe"),
                        "/PID " + process.Id + " /T /F") { UseShellExecute = false, CreateNoWindow = true })) stop.WaitForExit();
                    process.WaitForExit();
                    throw new IOException("转换超时，请检查文档是否过大或损坏。");
                }
                if (process.ExitCode != 0 || !File.Exists(produced)) throw new IOException("导出失败，请检查文件是否加密、损坏或被占用。");
            }
            using (var stream = File.OpenRead(produced)) {
                var signature = new byte[5];
                if (stream.Read(signature, 0, signature.Length) != 5 || Encoding.ASCII.GetString(signature) != "%PDF-")
                    throw new IOException("引擎未生成有效的 PDF。");
            }
        }
        static void ExportImage(string input, string target, string format) {
            using (var image = Image.FromFile(input)) {
                if (image.PropertyIdList.Contains(274)) {
                    int orientation = BitConverter.ToUInt16(image.GetPropertyItem(274).Value, 0);
                    var rotation = new[] { RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipNone,
                        RotateFlipType.RotateNoneFlipX, RotateFlipType.Rotate180FlipNone, RotateFlipType.Rotate180FlipX,
                        RotateFlipType.Rotate90FlipX, RotateFlipType.Rotate90FlipNone, RotateFlipType.Rotate270FlipX, RotateFlipType.Rotate270FlipNone };
                    if (orientation >= 2 && orientation <= 8) image.RotateFlip(rotation[orientation]);
                }
                if (format == "png") image.Save(target, ImageFormat.Png);
                else {
                    using (var canvas = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb)) {
                        using (var g = Graphics.FromImage(canvas)) { g.Clear(Color.White); g.DrawImageUnscaled(image, 0, 0); }
                        canvas.Save(target, ImageFormat.Jpeg);
                    }
                }
            }
        }
    }

    sealed class Window : Form {
        readonly ComboBox mode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 295 };
        readonly Label hint = new Label { Dock = DockStyle.Fill, AutoEllipsis=true, ForeColor=Color.FromArgb(61,79,91), TextAlign=ContentAlignment.MiddleLeft };
        readonly ListView queue = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, AllowDrop = true };
        readonly Label folder = new Label { Text = "与每个源文件相同的目录", AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
        readonly Label status = new Label { Text = "选择文件，或将文件拖入列表", AutoSize = true, Anchor = AnchorStyles.Left };
        readonly Button run = ButtonFor("开始转换", 112);
        readonly Button add = ButtonFor("添加文件", 98);
        readonly Button remove = ButtonFor("移除选中", 98);
        readonly Button browse = ButtonFor("更改目录", 98);
        string output;
        bool busy;
        string Format { get {
            return mode.SelectedItem == null ? null : ((KeyValuePair<string,string>)mode.SelectedItem).Key;
        } }
        static Button ButtonFor(string title, int width) {
            return new Button { Text = title, Width = width, Height = 32, Margin = new Padding(0, 0, 8, 0), UseVisualStyleBackColor = true };
        }
        static Label Caption(string title) {
            return new Label { Text = title, AutoSize = true, Anchor = AnchorStyles.Left };
        }
        public Window() {
            Text = "本地文件转换器";
            Font = new Font("Microsoft YaHei UI", 10);
            BackColor = Color.White;
            MinimumSize = new Size(700, 470);
            ClientSize = new Size(850, 510);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(22), ColumnCount = 1, RowCount = 7 };
            Controls.Add(layout);
            foreach (float h in new[] { 45f, 42f }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, h));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = "本地文件转换", Font = new Font(Font.FontFamily, 17, FontStyle.Bold),
                AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            var options = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(0, 6, 0, 0) };
            options.Controls.Add(Caption("目标格式  "));
            mode.DisplayMember = "Value";
            options.Controls.Add(mode);
            layout.Controls.Add(options, 0, 3);
            layout.Controls.Add(hint, 0, 4);
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            toolbar.Controls.Add(add);
            toolbar.Controls.Add(remove);
            layout.Controls.Add(toolbar, 0, 1);
            queue.Columns.Add("文件名", 360);
            queue.Columns.Add("状态", 370);
            layout.Controls.Add(queue, 0, 2);
            var destination = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3 };
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            destination.Controls.Add(Caption("保存到  "), 0, 0);
            destination.Controls.Add(folder, 1, 0);
            browse.Anchor = AnchorStyles.Right;
            destination.Controls.Add(browse, 2, 0);
            layout.Controls.Add(destination, 0, 5);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.Controls.Add(status, 0, 0);
            run.BackColor = Color.FromArgb(26, 85, 122);
            run.ForeColor = Color.White;
            run.FlatStyle = FlatStyle.Flat;
            run.Anchor = AnchorStyles.Right;
            footer.Controls.Add(run, 1, 0);
            layout.Controls.Add(footer, 0, 6);
            queue.ShowItemToolTips = true;
            queue.Resize += delegate { queue.Columns[0].Width = (int)(queue.ClientSize.Width * .43); queue.Columns[1].Width = Math.Max(140, queue.ClientSize.Width - queue.Columns[0].Width - 24); };
            add.Click += delegate {
                using (var picker = new OpenFileDialog { Multiselect = true, Filter = "支持的文件|" + string.Join(";", Converter.Inputs.Select(x => "*" + x)) }) {
                    if (picker.ShowDialog(this) == DialogResult.OK) AddFiles(picker.FileNames);
                }
            };
            queue.DragEnter += (s,e) => { if (!busy && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            queue.DragDrop += (s,e) => { if (!busy) AddFiles((string[])e.Data.GetData(DataFormats.FileDrop)); };
            remove.Click += delegate { foreach (var item in queue.SelectedItems.Cast<ListViewItem>().ToArray()) queue.Items.Remove(item); RefreshTargets(); };
            mode.SelectedIndexChanged += delegate { foreach (ListViewItem item in queue.Items) item.SubItems[1].Text = "等待转换"; RefreshStatus(); };
            browse.Click += delegate {
                using (var picker = new FolderBrowserDialog { Description = "请选择本地输出文件夹", SelectedPath = output ?? "" }) {
                    if (picker.ShowDialog(this) == DialogResult.OK) { output = picker.SelectedPath; folder.Text = output; }
                }
            };
            run.Click += async delegate { await RunBatch(); };
            FormClosing += (s,e) => { if (busy && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; status.Text = "正在转换，请在完成后关闭窗口。"; } };
            RefreshTargets();
        }
        void RefreshTargets() {
            string prior = Format;
            var paths = queue.Items.Cast<ListViewItem>().Select(i=>i.Name).ToArray();
            mode.BeginUpdate();
            mode.Items.Clear();
            if (paths.Length>0) {
                foreach (var target in Converter.Labels) {
                    if (paths.All(p=>Converter.Source(p)!=target.Key && Converter.Supports(p,target.Key)))
                        mode.Items.Add(target);
                }
            }
            int selected = mode.Items.Cast<KeyValuePair<string,string>>().ToList().FindIndex(p=>p.Key==prior);
            if (mode.Items.Count>0) mode.SelectedIndex=Math.Max(selected,0);
            mode.Enabled=mode.Items.Count>0;
            mode.EndUpdate();
            RefreshStatus();
        }
        void RefreshStatus() {
            int count = queue.Items.Cast<ListViewItem>().Count(i => Converter.Supports(i.Name, Format));
            status.Text = queue.Items.Count == 0 ? "选择文件，或将文件拖入列表" : "可转换 " + count + " 个文件";
            hint.Text = Converter.Describe(queue.Items.Cast<ListViewItem>().Select(i=>i.Name),Format);
            if (queue.Items.Count>0 && Format==null) hint.Text="本批文件没有共同的目标格式，请按类型分批转换。";
            run.Enabled = !busy && count > 0;
        }
        void AddFiles(IEnumerable<string> paths) {
            int skipped = 0;
            foreach (var path in paths) {
                if (!Converter.Inputs.Contains(Path.GetExtension(path),StringComparer.OrdinalIgnoreCase) || Directory.Exists(path)) { skipped++; continue; }
                if (queue.Items.ContainsKey(path)) continue;
                var item = queue.Items.Add(path, Path.GetFileName(path), -1);
                item.ToolTipText = path;
                item.SubItems.Add("等待转换");
            }
            RefreshTargets();
            if (skipped > 0) status.Text += "，已略过 " + skipped + " 个不支持的项目";
        }
        async Task RunBatch() {
            busy = true;
            add.Enabled = remove.Enabled = mode.Enabled = browse.Enabled = run.Enabled = false;
            string format = Format;
            var items = queue.Items.Cast<ListViewItem>().Where(i => Converter.Supports(i.Name, format)).ToArray();
            int finished = 0, failed = 0;
            try {
                foreach (var item in items) {
                    string path = item.Name, dest = output ?? Path.GetDirectoryName(path);
                    item.SubItems[1].Text = "转换中…";
                    status.Text = String.Format("正在转换 {0} / {1}", finished + failed + 1, items.Length);
                    try {
                        string result = await Task.Run(() => Converter.Run(path, dest, format));
                        item.SubItems[1].Text = "完成 · " + Path.GetFileName(result);
                        item.ToolTipText = result;
                        finished++;
                    } catch (Exception ex) { item.SubItems[1].Text = ex.Message; item.ToolTipText = ex.Message; failed++; }
                }
            } finally {
                busy = false;
                add.Enabled = remove.Enabled = mode.Enabled = browse.Enabled = true;
                RefreshStatus();
                status.Text = String.Format("完成 {0} 个，失败 {1} 个", finished, failed);
            }
        }
    }
    static class Program {
        [STAThread]
        static int Main(string[] args) {
            if (args.Length > 0) {
                if (args.Length != 4 || args[0] != "--convert" || !Converter.Labels.ContainsKey(args[1])) return 2;
                try { Console.WriteLine(Converter.Run(args[2], args[3], args[1])); return 0; }
                catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());
            return 0;
        }
    }
}
