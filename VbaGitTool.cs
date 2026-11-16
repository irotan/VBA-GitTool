// VBA Git 管理ツール 最終完成版 (ローカル管理特化カラーリング・実測サイズ完全レスポンシブ版)
// ビルド例 (C# 5 互換 / .NET Framework 4.x):
//   csc /target:winexe /out:VbaGitTool.exe VbaGitTool.cs /r:System.Windows.Forms.dll /r:System.Drawing.dll

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VbaGitTool
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main(string[] args)
        {
            try { SetProcessDPIAware(); } catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
            {
                MessageBox.Show("予期しないエラーが発生しました:\n" + e.Exception.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            string targetDir = Environment.CurrentDirectory;
            if (args != null && args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                string rawArg = args[0].Trim().Trim('"');
                if (Directory.Exists(rawArg)) targetDir = rawArg;
            }

            Application.Run(new MainForm(targetDir));
        }
    }

    // =========================================================================
    // プロセス実行 / 引数エスケープ
    // =========================================================================
    public class CmdResult
    {
        public int ExitCode;
        public string Output;   // stdout のみ
        public string Error;    // stderr のみ
        public bool Success { get { return ExitCode == 0; } }
    }

    public class CommandRunner
    {
        private readonly Action<string> _log;

        public CommandRunner(Action<string> log)
        {
            _log = log;
        }

        public static string Quote(string arg)
        {
            if (arg == null) arg = "";
            if (arg.Length > 0 && arg.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '"' }) < 0) return arg;

            var sb = new StringBuilder();
            sb.Append('"');
            int backslashes = 0;
            foreach (char c in arg)
            {
                if (c == '\\')
                {
                    backslashes++;
                }
                else if (c == '"')
                {
                    sb.Append('\\', backslashes * 2 + 1);
                    sb.Append('"');
                    backslashes = 0;
                }
                else
                {
                    if (backslashes > 0) { sb.Append('\\', backslashes); backslashes = 0; }
                    sb.Append(c);
                }
            }
            if (backslashes > 0) sb.Append('\\', backslashes * 2);
            sb.Append('"');
            return sb.ToString();
        }

        public static string JoinArgs(IEnumerable<string> args)
        {
            return string.Join(" ", args.Select(a => Quote(a)).ToArray());
        }

        public Task<CmdResult> GitAsync(string workDir, bool silent, params string[] args)
        {
            var all = new List<string> { "-c", "core.quotepath=false" };
            all.AddRange(args);
            return RunAsync("git", JoinArgs(all), workDir, silent);
        }

        public Task<CmdResult> RunAsync(string fileName, string arguments, string workDir, bool silent)
        {
            return Task.Run(() =>
            {
                bool isGit = fileName.Equals("git", StringComparison.OrdinalIgnoreCase);
                Encoding enc = isGit ? Encoding.UTF8 : Encoding.Default;

                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    WorkingDirectory = workDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = enc,
                    StandardErrorEncoding = enc
                };
                psi.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";

                var sbOut = new StringBuilder();
                var sbErr = new StringBuilder();
                object gate = new object();

                try
                {
                    using (var proc = new Process { StartInfo = psi })
                    {
                        proc.OutputDataReceived += (s, e) =>
                        {
                            if (e.Data == null) return;
                            lock (gate) { sbOut.AppendLine(e.Data); }
                            if (!silent) _log("  " + e.Data);
                        };
                        proc.ErrorDataReceived += (s, e) =>
                        {
                            if (e.Data == null) return;
                            lock (gate) { sbErr.AppendLine(e.Data); }
                            if (!silent) _log("  " + e.Data);
                        };

                        proc.Start();
                        proc.BeginOutputReadLine();
                        proc.BeginErrorReadLine();
                        proc.WaitForExit();

                        lock (gate)
                        {
                            return new CmdResult { ExitCode = proc.ExitCode, Output = sbOut.ToString(), Error = sbErr.ToString() };
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!silent) _log("[エラー] コマンドを実行できませんでした: " + fileName + " (" + ex.Message + ")");
                    return new CmdResult { ExitCode = -1, Output = "", Error = ex.Message };
                }
            });
        }
    }

    // =========================================================================
    // 入力検証
    // =========================================================================
    public static class Validation
    {
        private static readonly Regex HashRegex = new Regex("^[0-9a-fA-F]{4,40}$", RegexOptions.Compiled);

        public static bool IsValidCommitHash(string s)
        {
            return !string.IsNullOrEmpty(s) && HashRegex.IsMatch(s);
        }

        public static bool IsValidFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            name = name.Trim();
            if (name == "." || name == "..") return false;
            if (name.EndsWith(".") || name.EndsWith(" ")) return false;
            return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;
        }

        public static bool IsSafeCloneUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;
            url = url.Trim();
            if (url.StartsWith("-")) return false;
            if (url.Any(c => char.IsWhiteSpace(c) || char.IsControl(c))) return false;
            return true;
        }

        public static string SanitizeMessage(string msg)
        {
            if (msg == null) return "";
            return msg.Replace("\r", " ").Replace("\n", " ").Trim();
        }
    }

    // =========================================================================
    // 右クリックメニュー登録 (HKCU / 管理者権限不要)
    // =========================================================================
    public static class ContextMenuRegistrar
    {
        private const string RegKeyBg = @"Software\Classes\Directory\Background\shell\VbaGitTool";
        private const string RegKeyDir = @"Software\Classes\Directory\shell\VbaGitTool";

        public static bool IsRegistered()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RegKeyBg))
            {
                return key != null;
            }
        }

        public static void Unregister()
        {
            Registry.CurrentUser.DeleteSubKeyTree(RegKeyBg, false);
            Registry.CurrentUser.DeleteSubKeyTree(RegKeyDir, false);
        }

        public static void Register(string exePath)
        {
            Write(RegKeyBg, exePath, "\"" + exePath + "\" \"%V\"");
            Write(RegKeyDir, exePath, "\"" + exePath + "\" \"%1\"");
        }

        private static void Write(string keyPath, string exePath, string command)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                key.SetValue("", "VBA Git 管理ツールを開く");
                key.SetValue("Icon", exePath);
                using (var cmd = key.CreateSubKey("command"))
                {
                    cmd.SetValue("", command);
                }
            }
        }
    }

    // =========================================================================
    // 退避（バックアップ）/ .gitignore 補助 / フォルダ自動補正
    // =========================================================================
    public static class ProjectFiles
    {
        public static readonly string[] DefaultIgnoreLines = { "bin/*", "!bin/.gitkeep", "~$*" };
        private const int KeepBackups = 10;

        public static string ResolveProjectRoot(string dir, out bool adjusted)
        {
            adjusted = false;
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)) return dir;

            string trimmed = dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (HasGitFolder(trimmed)) return trimmed;

            DirectoryInfo parentInfo = null;
            try { parentInfo = Directory.GetParent(trimmed); } catch { }
            if (parentInfo == null || !parentInfo.Exists) return trimmed;

            string parentPath = parentInfo.FullName;
            string currentFolderName = Path.GetFileName(trimmed);

            bool parentHasGit = HasGitFolder(parentPath);
            bool isSubFolderName = currentFolderName.Equals("bin", StringComparison.OrdinalIgnoreCase)
                                || currentFolderName.Equals("src", StringComparison.OrdinalIgnoreCase);
            bool parentHasMarkers = Directory.Exists(Path.Combine(parentPath, "bin"))
                                 || Directory.Exists(Path.Combine(parentPath, "src"));

            if (parentHasGit || (isSubFolderName && parentHasMarkers))
            {
                adjusted = true;
                return parentPath;
            }

            return trimmed;
        }

        private static bool HasGitFolder(string path)
        {
            string dotGit = Path.Combine(path, ".git");
            return Directory.Exists(dotGit) || File.Exists(dotGit);
        }

        public static string BackupWorkingState(string dir, Action<string> log)
        {
            try
            {
                string src = Path.Combine(dir, "src");
                string bin = Path.Combine(dir, "bin");
                if (!Directory.Exists(src) && !Directory.Exists(bin)) return "";

                string gitDir = Path.Combine(dir, ".git");
                string root = Directory.Exists(gitDir)
                    ? Path.Combine(gitDir, "vbagit_backup")
                    : Path.Combine(Path.GetTempPath(), "vbagit_backup_" + Path.GetFileName(dir.TrimEnd('\\', '/')));

                string dest = Path.Combine(root, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                Directory.CreateDirectory(dest);

                if (Directory.Exists(src)) CopyDirectory(src, Path.Combine(dest, "src"));
                if (Directory.Exists(bin)) CopyDirectory(bin, Path.Combine(dest, "bin"));

                foreach (string old in Directory.GetDirectories(root).OrderByDescending(d => d).Skip(KeepBackups))
                {
                    try { Directory.Delete(old, true); } catch { }
                }
                return dest;
            }
            catch (Exception ex)
            {
                if (log != null) log("[警告] 退避に失敗しました: " + ex.Message);
                return null;
            }
        }

        private static void CopyDirectory(string from, string to)
        {
            Directory.CreateDirectory(to);
            foreach (string f in Directory.GetFiles(from))
            {
                string name = Path.GetFileName(f);
                if (name.StartsWith("~$")) continue;
                File.Copy(f, Path.Combine(to, name), true);
            }
            foreach (string d in Directory.GetDirectories(from))
            {
                CopyDirectory(d, Path.Combine(to, Path.GetFileName(d)));
            }
        }

        public static void EnsureDefaultGitignore(string dir)
        {
            string path = Path.Combine(dir, ".gitignore");
            var lines = new List<string>();
            if (File.Exists(path))
            {
                lines.AddRange(File.ReadAllLines(path, Encoding.UTF8).Select(l => l.TrimEnd()));
            }
            foreach (string item in DefaultIgnoreLines)
            {
                if (!lines.Contains(item)) lines.Add(item);
            }
            File.WriteAllText(path, string.Join("\r\n", lines.Where(l => l.Length > 0).ToArray()) + "\r\n", new UTF8Encoding(false));

            string bin = Path.Combine(dir, "bin");
            Directory.CreateDirectory(bin);
            string keep = Path.Combine(bin, ".gitkeep");
            if (!File.Exists(keep)) File.WriteAllText(keep, "", new UTF8Encoding(false));
        }

        public static bool GitignoreCoversBin(string dir)
        {
            string path = Path.Combine(dir, ".gitignore");
            if (!File.Exists(path)) return false;
            var accepted = new[] { "bin", "bin/", "bin/*", "/bin", "/bin/", "/bin/*" };
            return File.ReadAllLines(path, Encoding.UTF8).Any(l => accepted.Contains(l.Trim()));
        }
    }

    // =========================================================================
    // メインフォーム (ローカル管理特化カラーリング ＋ 実測フォント基準レスポンシブ設計)
    // =========================================================================
    public class MainForm : Form
    {
        private string _targetDir;
        private readonly bool _initialDirAdjusted;
        private readonly string _appDir;
        private readonly string _masterVbacPath;
        private readonly CommandRunner _cmd;

        // カラー定義（ローカル操作のみを強調し、それ以外は標準の白背景・黒文字に統一）
        private static readonly Color BgBlue  = Color.FromArgb(230, 242, 255);
        private static readonly Color FgBlue  = Color.FromArgb(0, 70, 160);
        private static readonly Color BgGreen = Color.FromArgb(215, 245, 225);
        private static readonly Color FgGreen = Color.FromArgb(20, 100, 40);
        private static readonly Color BgWhite = Color.White;
        private static readonly Color FgBlack = Color.Black;

        // 状態
        private bool _busy;
        private bool _gitAvailable = true;
        private string _branch;
        private bool _hasRemote;

        // UI
        private Panel pnlTop;
        private Label lblTargetDir;
        private Label lblGitInfo;
        private Label lblMsg;
        private Label lblTag;
        private Label lblHist;
        private Label lblLog;
        private TextBox txtCommitMsg;
        private TextBox txtTagName;
        private Button btnSave;
        private Button btnSaveAndPush;
        private Button btnPushOnly;
        private Button btnPull;
        private Button btnRestoreLatest;
        private ListBox lstHistory;
        private Button btnRestoreLocal;
        private Button btnRestoreGithub;
        private Button btnAddTag;
        private Button btnRefreshHist;
        private Button btnInitLocalNew;
        private Button btnCloneNew;
        private Button btnInitProject;
        private Button btnToggleReg;
        private Button btnChangeDir;
        private Button btnHelp;
        private GroupBox grpSave;
        private GroupBox grpRestore;
        private GroupBox grpSetup;
        private TextBox txtLog;
        private HelpForm _helpForm;

        private class HistoryItem
        {
            public string Hash;
            public string Display;
            public override string ToString() { return Display; }
        }

        public MainForm(string initialDir)
        {
            bool adjusted;
            _targetDir = ProjectFiles.ResolveProjectRoot(initialDir, out adjusted);
            _initialDirAdjusted = adjusted;

            _appDir = AppDomain.CurrentDomain.BaseDirectory;
            _masterVbacPath = Path.Combine(_appDir, "vbac.wsf");
            _cmd = new CommandRunner(AppendLog);

            InitializeComponents();
            this.Layout += (s, e) => ApplyResponsiveLayout();
            this.Shown += async (s, e) =>
            {
                ApplyResponsiveLayout();
                await InitialLoadAsync();
            };
            this.FormClosing += (s, e) =>
            {
                if (_busy)
                {
                    MessageBox.Show("処理の実行中です。完了するまでお待ちください。", "終了できません",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                }
            };
        }

        /// <summary>
        /// 指定した背景色・文字色・太字有無でボタンを生成する。
        /// 無効化(Enabled=false)時にも指定した文字色（青・緑・黒）がグレーに潰れず視認できるよう描画を補助する。
        /// </summary>
        private Button CreateStyledButton(string text, Color backColor, Color foreColor, bool bold)
        {
            var btn = new Button
            {
                Text = text,
                BackColor = backColor,
                ForeColor = foreColor,
                UseVisualStyleBackColor = false,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Yu Gothic UI", 9.5f, bold ? FontStyle.Bold : FontStyle.Regular),
                AutoEllipsis = false
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 205, 210);
            btn.FlatAppearance.BorderSize = 1;

            // Enabled=false のとき標準描画だと文字色が強制的に灰色になり指定色が見えなくなるため、無効時のみ指定色で文字を上書き描画
            btn.Paint += (s, e) =>
            {
                var b = (Button)s;
                if (!b.Enabled)
                {
                    var rect = new Rectangle(1, 1, b.Width - 2, b.Height - 2);
                    using (var brush = new SolidBrush(b.BackColor))
                    {
                        e.Graphics.FillRectangle(brush, rect);
                    }
                    TextRenderer.DrawText(
                        e.Graphics,
                        b.Text,
                        b.Font,
                        rect,
                        b.ForeColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
                }
            };
            return btn;
        }

        private static int MeasureButtonWidth(Control btn, int extraPadding)
        {
            Size sz = TextRenderer.MeasureText(btn.Text, btn.Font);
            return sz.Width + extraPadding;
        }

        // ---------------------------------------------------------------------
        // UI 構築 (AutoScaleMode.None で勝手な縮小・拡大を防止)
        // ---------------------------------------------------------------------
        private void InitializeComponents()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            this.Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular);

            int fh = this.Font.Height;
            int initW = Math.Max(940, fh * 52);
            int initH = Math.Max(740, fh * 41);

            Rectangle workArea = Screen.PrimaryScreen.WorkingArea;
            if (initW > workArea.Width - 20) initW = workArea.Width - 20;
            if (initH > workArea.Height - 20) initH = workArea.Height - 20;

            this.Text = "VBA Git 管理ツール";
            this.ClientSize = new Size(initW, initH);
            this.MinimumSize = new Size(Math.Min(860, initW), Math.Min(680, initH));
            this.StartPosition = FormStartPosition.CenterScreen;

            // 上部ヘッダー
            pnlTop = new Panel { BackColor = Color.FromArgb(240, 244, 248) };
            lblTargetDir = new Label { AutoSize = false, Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            lblGitInfo = new Label { AutoSize = false, ForeColor = Color.FromArgb(40, 80, 140), AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };

            btnChangeDir = CreateStyledButton("フォルダ変更...", BgWhite, FgBlack, false);
            btnChangeDir.Click += (s, e) => ChangeTargetDirectory();

            btnHelp = CreateStyledButton("ヘルプ / 使い方", BgBlue, FgBlue, false);
            btnHelp.Click += (s, e) => ShowHelpWindow();

            pnlTop.Controls.AddRange(new Control[] { lblTargetDir, lblGitInfo, btnChangeDir, btnHelp });

            // Group 1: 日常のセーブ・GitHub同期
            grpSave = new GroupBox { Text = "1. 日常のセーブ・GitHub同期" };
            lblMsg = new Label { Text = "コミットメモ:", AutoSize = true };
            txtCommitMsg = new TextBox();
            lblTag = new Label { Text = "タグ(任意):", AutoSize = true };
            txtTagName = new TextBox();

            // ローカル保存のみ青＆太字、GitHub送信系は標準（白背景・黒文字・太字なし）
            btnSave = CreateStyledButton("Save (ローカル保存のみ)", BgBlue, FgBlue, true);
            btnSave.Click += async (s, e) => await RunOperationAsync(() => ExecuteSaveAsync(false));

            btnSaveAndPush = CreateStyledButton("Save ＆ GitHub送信 (一括)", BgWhite, FgBlack, false);
            btnSaveAndPush.Click += async (s, e) => await RunOperationAsync(() => ExecuteSaveAsync(true));

            btnPushOnly = CreateStyledButton("GitHub送信のみ (Push)", BgWhite, FgBlack, false);
            btnPushOnly.Click += async (s, e) => await RunOperationAsync(() => ExecutePushOnlyAsync());

            grpSave.Controls.AddRange(new Control[] { lblMsg, txtCommitMsg, lblTag, txtTagName, btnSave, btnSaveAndPush, btnPushOnly });

            // Group 2: 取得・復元（やり直し）
            grpRestore = new GroupBox { Text = "2. 取得・復元（やり直し） ※Excelを閉じてから実行" };
            lblHist = new Label { Text = "過去の履歴・タグ一覧（選択して復元 or タグ付け）:", AutoSize = true };

            // ⑦「履歴再読込」ボタンを青にする
            btnRefreshHist = CreateStyledButton("履歴再読込", BgBlue, FgBlue, false);
            btnRefreshHist.Click += async (s, e) => await RunOperationAsync(() => RefreshHistoryListAsync());

            lstHistory = new ListBox { Font = new Font("Consolas", 9.5f), HorizontalScrollbar = true, IntegralHeight = false };

            // 選択した履歴へ復元 (ローカル) は青＆太字
            btnRestoreLocal = CreateStyledButton("選択した履歴へ復元 (ローカル)", BgBlue, FgBlue, true);
            btnRestoreLocal.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreSelectedAsync(false));

            // ②「直前のSave状態へ戻す (ローカル)」にして、緑の塗りつぶしと文字色を緑にする
            btnRestoreLatest = CreateStyledButton("直前のSave状態へ戻す (ローカル)", BgGreen, FgGreen, false);
            btnRestoreLatest.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreLatestAsync());

            // ③ Pull含め下段のボタンはすべて標準（白背景・黒文字・太字なし）
            btnPull = CreateStyledButton("Pull (GitHub最新取得＋Excel反映)", BgWhite, FgBlack, false);
            btnPull.Click += async (s, e) => await RunOperationAsync(() => ExecutePullAsync());

            btnRestoreGithub = CreateStyledButton("選択した履歴へ復元 (GitHub取得後)", BgWhite, FgBlack, false);
            btnRestoreGithub.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreSelectedAsync(true));

            btnAddTag = CreateStyledButton("選択履歴にタグ付与", BgWhite, FgBlack, false);
            btnAddTag.Click += async (s, e) => await RunOperationAsync(() => ExecuteTagSelectedAsync());

            grpRestore.Controls.AddRange(new Control[] { lblHist, btnRefreshHist, lstHistory, btnRestoreLocal, btnRestoreLatest, btnPull, btnRestoreGithub, btnAddTag });

            // Group 3: 新規構築・初期設定・メンテナンス
            grpSetup = new GroupBox { Text = "3. 新規構築・初期設定・メンテナンス" };

            btnInitLocalNew = CreateStyledButton("新規作成 (ローカル専用)", BgWhite, FgBlack, false);
            btnInitLocalNew.Click += async (s, e) => await RunOperationAsync(() => ExecuteInitLocalNewAsync());

            btnCloneNew = CreateStyledButton("新規クローン (GitHubから)", BgWhite, FgBlack, false);
            btnCloneNew.Click += async (s, e) => await RunOperationAsync(() => ExecuteCloneNewAsync());

            // ④ 初期設定ボタンの青の塗りつぶしをやめて、文字色も標準の黒（太字なし）にする
            btnInitProject = CreateStyledButton("初期設定 (.gitignore / 監視解除)", BgWhite, FgBlack, false);
            btnInitProject.Click += async (s, e) => await RunOperationAsync(() => OpenInitProjectDialogAsync());

            grpSetup.Controls.AddRange(new Control[] { btnInitLocalNew, btnCloneNew, btnInitProject });

            // 4. 実行ログ
            lblLog = new Label { Text = "実行ログ:", AutoSize = true };
            // ⑤ 右クリックメニュー登録ボタンの緑の塗りつぶしをやめて、文字色は標準の黒（太字なし）にする
            btnToggleReg = CreateStyledButton("右クリックメニュー登録済み (解除する)", BgWhite, FgBlack, false);
            btnToggleReg.Click += (s, e) => ToggleContextMenuRegistration();

            txtLog = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(220, 235, 220),
                Font = new Font("Consolas", 9.5f)
            };

            this.Controls.AddRange(new Control[] { pnlTop, grpSave, grpRestore, grpSetup, lblLog, btnToggleReg, txtLog });
            ApplyResponsiveLayout();
        }

        /// <summary>
        /// ウィンドウサイズと実測文字サイズに基づき、すべてのコントロールを重なり・欠け無しで配置する。
        /// </summary>
        private void ApplyResponsiveLayout()
        {
            if (pnlTop == null || this.ClientSize.Width < 200 || this.ClientSize.Height < 200) return;

            int cw = this.ClientSize.Width;
            int ch = this.ClientSize.Height;

            int fh = this.Font.Height;
            int btnH = fh + 16;
            int smallBtnH = fh + 10;
            int pad = 12;
            int innerPad = 14;
            int gap = 10;

            // 1. 上部ヘッダー
            int topH = fh * 2 + 24;
            pnlTop.Bounds = new Rectangle(0, 0, cw, topH);

            int helpW = MeasureButtonWidth(btnHelp, 28);
            int chgW = MeasureButtonWidth(btnChangeDir, 28);
            int topBtnY = (topH - btnH) / 2;

            btnHelp.Bounds = new Rectangle(cw - pad - helpW, topBtnY, helpW, btnH);
            btnChangeDir.Bounds = new Rectangle(btnHelp.Left - 8 - chgW, topBtnY, chgW, btnH);

            int infoW = Math.Max(100, btnChangeDir.Left - pad - 12);
            lblTargetDir.Bounds = new Rectangle(pad, 8, infoW, fh + 4);
            lblGitInfo.Bounds = new Rectangle(pad, 12 + fh, infoW, fh + 4);

            int grpW = cw - pad * 2;
            int innerW = grpW - innerPad * 2;

            // 2. Group 1 (Save)
            int grp1Y = topH + 8;
            int grpTopMargin = fh + 8;
            int txtH = txtCommitMsg.PreferredHeight;

            int lblMsgW = TextRenderer.MeasureText(lblMsg.Text, lblMsg.Font).Width;
            int lblTagW = TextRenderer.MeasureText(lblTag.Text, lblTag.Font).Width;

            lblMsg.Location = new Point(innerPad, grpTopMargin + (txtH - fh) / 2);
            int tagW = Math.Max(130, innerW / 6);
            txtTagName.Bounds = new Rectangle(grpW - innerPad - tagW, grpTopMargin, tagW, txtH);
            lblTag.Location = new Point(txtTagName.Left - lblTagW - 6, grpTopMargin + (txtH - fh) / 2);

            int msgLeft = innerPad + lblMsgW + 6;
            int msgW = Math.Max(100, lblTag.Left - 16 - msgLeft);
            txtCommitMsg.Bounds = new Rectangle(msgLeft, grpTopMargin, msgW, txtH);

            int saveBtnY = grpTopMargin + txtH + 10;
            int w3_1 = (innerW - gap * 2) / 3;
            int w3_2 = (innerW - gap * 2) / 3;
            int w3_3 = innerW - gap * 2 - w3_1 - w3_2;

            btnSave.Bounds = new Rectangle(innerPad, saveBtnY, w3_1, btnH);
            btnSaveAndPush.Bounds = new Rectangle(btnSave.Right + gap, saveBtnY, w3_2, btnH);
            btnPushOnly.Bounds = new Rectangle(btnSaveAndPush.Right + gap, saveBtnY, w3_3, btnH);

            int grp1H = saveBtnY + btnH + 12;
            grpSave.Bounds = new Rectangle(pad, grp1Y, grpW, grp1H);

            // 3. Group 3 (Setup) の高さを確定
            int setupBtnY = fh + 8;
            int grp3H = setupBtnY + btnH + 12;

            // 4. 残り高さを Group 2 (Restore) と Log に配分 (58% : 42%)
            int logHeaderH = smallBtnH + 6;
            int usedFixedH = grp1Y + grp1H + 8 + grp3H + 8 + logHeaderH + pad;
            int availFlexH = Math.Max(220, ch - usedFixedH);
            int grp2H = Math.Max(fh * 10, (int)(availFlexH * 0.58));

            int grp2Y = grpSave.Bottom + 8;
            grpRestore.Bounds = new Rectangle(pad, grp2Y, grpW, grp2H);

            // Group 2 内部
            int rHeaderY = fh + 6;
            int refW = MeasureButtonWidth(btnRefreshHist, 24);
            btnRefreshHist.Bounds = new Rectangle(grpW - innerPad - refW, rHeaderY, refW, smallBtnH);
            lblHist.Location = new Point(innerPad, rHeaderY + (smallBtnH - fh) / 2);

            int rBotRowY = grp2H - 12 - btnH;
            int rMidRowY = rBotRowY - 8 - btnH;

            // 中段：2等分 (ローカル復元)
            int w2_1 = (innerW - gap) / 2;
            int w2_2 = innerW - gap - w2_1;
            btnRestoreLocal.Bounds = new Rectangle(innerPad, rMidRowY, w2_1, btnH);
            btnRestoreLatest.Bounds = new Rectangle(btnRestoreLocal.Right + gap, rMidRowY, w2_2, btnH);

            // 下段：3ボタン (37% : 37% : 26% で配分)
            int subAvailW = innerW - gap * 2;
            int ws1 = (int)(subAvailW * 0.37);
            int ws2 = (int)(subAvailW * 0.37);
            int ws3 = subAvailW - ws1 - ws2;
            btnPull.Bounds = new Rectangle(innerPad, rBotRowY, ws1, btnH);
            btnRestoreGithub.Bounds = new Rectangle(btnPull.Right + gap, rBotRowY, ws2, btnH);
            btnAddTag.Bounds = new Rectangle(btnRestoreGithub.Right + gap, rBotRowY, ws3, btnH);

            // 履歴リストボックス
            int lstTop = btnRefreshHist.Bottom + 6;
            int lstH = Math.Max(40, rMidRowY - 8 - lstTop);
            lstHistory.Bounds = new Rectangle(innerPad, lstTop, innerW, lstH);

            // 5. Group 3 (Setup)
            int grp3Y = grpRestore.Bottom + 8;
            grpSetup.Bounds = new Rectangle(pad, grp3Y, grpW, grp3H);
            btnInitLocalNew.Bounds = new Rectangle(innerPad, setupBtnY, w3_1, btnH);
            btnCloneNew.Bounds = new Rectangle(btnInitLocalNew.Right + gap, setupBtnY, w3_2, btnH);
            btnInitProject.Bounds = new Rectangle(btnCloneNew.Right + gap, setupBtnY, w3_3, btnH);

            // 6. 実行ログエリア
            int logHeaderY = grpSetup.Bottom + 8;
            int regW = MeasureButtonWidth(btnToggleReg, 28);
            btnToggleReg.Bounds = new Rectangle(cw - pad - regW, logHeaderY, regW, smallBtnH);
            lblLog.Location = new Point(pad + 2, logHeaderY + (smallBtnH - fh) / 2);

            int txtLogTop = btnToggleReg.Bottom + 6;
            int txtLogH = Math.Max(60, ch - pad - txtLogTop);
            txtLog.Bounds = new Rectangle(pad, txtLogTop, grpW, txtLogH);
        }

        private void ShowHelpWindow()
        {
            if (_helpForm == null || _helpForm.IsDisposed)
            {
                _helpForm = new HelpForm();
                _helpForm.Show(this);
            }
            else
            {
                _helpForm.BringToFront();
            }
        }

        // ---------------------------------------------------------------------
        // 実行制御（排他・例外処理）
        // ---------------------------------------------------------------------
        private async Task RunOperationAsync(Func<Task> body)
        {
            if (_busy) return;
            SetBusy(true);

            Exception error = null;
            try
            {
                await body();
            }
            catch (Exception ex)
            {
                error = ex;
            }

            try
            {
                await LoadDirectoryContextAsync();
            }
            catch (Exception ex2)
            {
                AppendLog("[例外] 状態の再読込に失敗: " + ex2.Message);
            }

            SetBusy(false);

            if (error != null)
            {
                AppendLog("[例外] " + error.Message);
                MessageBox.Show("処理中にエラーが発生しました:\n" + error.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            bool enabled = !busy;
            grpSave.Enabled = enabled;
            grpRestore.Enabled = enabled;
            grpSetup.Enabled = enabled;
            btnChangeDir.Enabled = enabled;
            btnToggleReg.Enabled = enabled;
            this.UseWaitCursor = busy;
        }

        private async Task InitialLoadAsync()
        {
            SetBusy(true);
            try
            {
                if (_initialDirAdjusted)
                {
                    AppendLog("[案内] サブフォルダ(bin/src等)から起動されたため、1階層上の親フォルダへ自動補正しました: " + _targetDir);
                }

                var r = await _cmd.RunAsync("git", "--version", _appDir, true);
                _gitAvailable = r.Success;
                if (!_gitAvailable)
                {
                    MessageBox.Show("git が見つかりません。\nGit for Windows をインストールし、PATH が通っていることを確認してください。",
                        "git 未検出", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendLog("[エラー] git を実行できません。 " + r.Error.Trim());
                }
                await LoadDirectoryContextAsync();
            }
            catch (Exception ex)
            {
                AppendLog("[例外] 初期化に失敗: " + ex.Message);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task LoadDirectoryContextAsync()
        {
            bool adjusted;
            string resolved = ProjectFiles.ResolveProjectRoot(_targetDir, out adjusted);
            if (adjusted)
            {
                _targetDir = resolved;
                AppendLog("[案内] 親フォルダ（開発用フォルダ）を検出したため、対象フォルダを自動補正しました: " + _targetDir);
            }

            lblTargetDir.Text = "対象フォルダ: " + _targetDir;
            UpdateRegButtonText();

            if (!_gitAvailable)
            {
                lblGitInfo.Text = "状態: git が見つかりません";
                SetProjectButtonsEnabled(false, false, false);
                btnInitLocalNew.Enabled = false;
                btnCloneNew.Enabled = false;
                lstHistory.Items.Clear();
                return;
            }

            if (!File.Exists(_masterVbacPath))
            {
                AppendLog("[警告] アプリと同じフォルダに vbac.wsf が見つかりません: " + _masterVbacPath);
            }

            string dotGit = Path.Combine(_targetDir, ".git");
            bool isGitRepo = Directory.Exists(dotGit) || File.Exists(dotGit);
            if (!isGitRepo)
            {
                lblGitInfo.Text = "状態: Git未初期化フォルダ（親フォルダとして「新規作成」「クローン」が可能です）";
                _branch = null;
                _hasRemote = false;
                SetProjectButtonsEnabled(false, false, false);
                lstHistory.Items.Clear();
                return;
            }

            var br = await _cmd.GitAsync(_targetDir, true, "symbolic-ref", "--short", "-q", "HEAD");
            _branch = br.Success ? br.Output.Trim() : null;
            if (string.IsNullOrEmpty(_branch)) _branch = null;

            var rm = await _cmd.GitAsync(_targetDir, true, "remote", "get-url", "origin");
            string remoteUrl = rm.Success ? rm.Output.Trim() : "";
            _hasRemote = !string.IsNullOrEmpty(remoteUrl);

            string branchText = _branch ?? "(detached HEAD)";
            string remoteText = _hasRemote ? remoteUrl : "(リモート未設定・ローカル専用)";
            lblGitInfo.Text = string.Format("ブランチ: {0}   |   接続先(origin): {1}", branchText, remoteText);

            bool canSync = _hasRemote && _branch != null;
            SetProjectButtonsEnabled(true, canSync, _hasRemote);

            if (_branch == null)
            {
                AppendLog("[警告] ブランチ上にいません (detached HEAD)。GitHub送信/取得は無効化しています。");
            }

            await RefreshHistoryListAsync();
        }

        private void SetProjectButtonsEnabled(bool isGitRepo, bool canSync, bool hasRemote)
        {
            btnInitLocalNew.Enabled = !isGitRepo;
            btnCloneNew.Enabled = !isGitRepo;

            btnSave.Enabled = isGitRepo;
            btnRestoreLatest.Enabled = isGitRepo;
            btnRestoreLocal.Enabled = isGitRepo;
            btnAddTag.Enabled = isGitRepo;
            btnRefreshHist.Enabled = isGitRepo;
            btnInitProject.Enabled = isGitRepo;

            btnSaveAndPush.Enabled = isGitRepo && canSync;
            btnPushOnly.Enabled = isGitRepo && canSync;
            btnPull.Enabled = isGitRepo && canSync;
            btnRestoreGithub.Enabled = isGitRepo && hasRemote;

            this.Invalidate(true);
        }

        private void ChangeTargetDirectory()
        {
            if (_busy) return;
            using (var dlg = new FolderBrowserDialog { Description = "操作対象のフォルダを選択してください", SelectedPath = _targetDir })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    bool adjusted;
                    _targetDir = ProjectFiles.ResolveProjectRoot(dlg.SelectedPath, out adjusted);
                    txtLog.Clear();
                    if (adjusted)
                    {
                        AppendLog("[案内] サブフォルダが選択されたため、1階層上の親フォルダへ自動補正しました: " + _targetDir);
                    }
                    var ignored = RunOperationAsync(() => Task.FromResult(0));
                }
            }
        }

        // ---------------------------------------------------------------------
        // 安全装置 / vbac
        // ---------------------------------------------------------------------
        private bool IsExcelOpenInBin(string dir)
        {
            string binDir = Path.Combine(dir, "bin");
            if (!Directory.Exists(binDir)) return false;

            var lockFiles = Directory.GetFiles(binDir, "~$*");
            if (lockFiles.Length > 0)
            {
                string fileNames = string.Join("\n", lockFiles.Select(f => Path.GetFileName(f)).ToArray());
                MessageBox.Show(
                    "binフォルダ内のExcelブックが開かれています！\n\n検出された一時ファイル:\n" + fileNames + "\n\nExcelを完全に閉じてから再実行してください。",
                    "Excel使用中警告（安全装置）",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return true;
            }
            return false;
        }

        private async Task<bool> RunVbacAsync(string dir, string mode)
        {
            if (!File.Exists(_masterVbacPath))
            {
                MessageBox.Show("アプリフォルダ内に vbac.wsf が存在しません。\n配置先: " + _masterVbacPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string binDir = Path.Combine(dir, "bin");
            if (!Directory.Exists(binDir)) Directory.CreateDirectory(binDir);

            string tempVbac = Path.Combine(dir, "vbac.wsf");
            bool copiedTemp = false;
            try
            {
                if (!File.Exists(tempVbac))
                {
                    File.Copy(_masterVbacPath, tempVbac, true);
                    copiedTemp = true;
                }

                AppendLog(string.Format("> cscript //nologo vbac.wsf {0}", mode));
                var res = await _cmd.RunAsync("cscript", "//nologo vbac.wsf " + mode, dir, false);
                return res.ExitCode == 0;
            }
            finally
            {
                if (copiedTemp && File.Exists(tempVbac))
                {
                    try { File.Delete(tempVbac); } catch { }
                }
            }
        }

        private async Task<bool> BackupBeforeOverwriteAsync(string dir)
        {
            string path = await Task.Run(() => ProjectFiles.BackupWorkingState(dir, AppendLog));
            if (path == null)
            {
                return MessageBox.Show("現在の状態の退避に失敗しました。\n退避なしで続行すると、未保存の編集内容を失う可能性があります。\n続行しますか？",
                    "退避失敗", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes;
            }
            if (path.Length > 0) AppendLog("・安全のため現在の状態を退避しました: " + path);
            return true;
        }

        private async Task<bool> IsValidTagNameAsync(string dir, string tag)
        {
            if (string.IsNullOrWhiteSpace(tag) || tag.StartsWith("-")) return false;
            var r = await _cmd.GitAsync(dir, true, "check-ref-format", "refs/tags/" + tag);
            return r.Success;
        }

        private async Task<bool> HasStagedChangesAsync(string dir)
        {
            var r = await _cmd.GitAsync(dir, true, "diff", "--cached", "--quiet");
            return r.ExitCode != 0;
        }

        // ---------------------------------------------------------------------
        // 日常操作
        // ---------------------------------------------------------------------
        private bool ConfirmGitignoreBeforeSave(string dir)
        {
            if (ProjectFiles.GitignoreCoversBin(dir)) return true;

            var r = MessageBox.Show(
                ".gitignore に bin/ の除外設定が見つかりません。\nこのまま Save すると、binフォルダ内のExcelブック(.xlsm等)がそのままGitに登録される可能性があります。\n\n" +
                "[はい] 標準の除外設定（bin/* , ~$* ）を自動追加して続行\n" +
                "[いいえ] 除外設定なしで続行\n" +
                "[キャンセル] Saveを中止",
                "除外設定の確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.Yes)
            {
                ProjectFiles.EnsureDefaultGitignore(dir);
                AppendLog("・.gitignore に標準の除外設定を追加しました。");
                AppendLog("  ※すでにbin内のファイルをGit登録済みの場合は「初期設定」で監視解除してください。");
            }
            return true;
        }

        private async Task ExecuteSaveAsync(bool alsoPush)
        {
            string dir = _targetDir;

            if (!ConfirmGitignoreBeforeSave(dir)) return;

            string tag = txtTagName.Text.Trim();
            if (tag.Length > 0 && !await IsValidTagNameAsync(dir, tag))
            {
                MessageBox.Show("タグ名として使用できない文字が含まれています。\n（空白、~ ^ : ? * [ \\ などは使用できません）", "タグ名エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (alsoPush && (!_hasRemote || _branch == null))
            {
                AppendLog("[エラー] リモート未設定、またはブランチ上にいないため GitHub へ送信できません。");
                return;
            }

            AppendLog("========== Save 処理開始 ==========");
            if (!await RunVbacAsync(dir, "decombine"))
            {
                AppendLog("[エラー] vbac.wsf decombine に失敗しました。");
                return;
            }

            var addRes = await _cmd.GitAsync(dir, false, "add", "-A");
            if (!addRes.Success)
            {
                AppendLog("[エラー] git add に失敗しました。");
                return;
            }

            string msg = Validation.SanitizeMessage(txtCommitMsg.Text);
            if (msg.Length == 0) msg = "作業セーブ (" + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ")";

            if (await HasStagedChangesAsync(dir))
            {
                var commitRes = await _cmd.GitAsync(dir, false, "commit", "-m", msg);
                if (!commitRes.Success)
                {
                    AppendLog("[エラー] コミットに失敗しました。user.name / user.email の設定、またはフックのエラーを確認してください。");
                    return;
                }
                txtCommitMsg.Clear();
            }
            else
            {
                AppendLog("[案内] 前回セーブからコードの変更点がないため、新規コミットはスキップされました。");
            }

            if (tag.Length > 0)
            {
                var tagRes = await _cmd.GitAsync(dir, false, "tag", "-a", tag, "-m", msg);
                if (tagRes.Success) txtTagName.Clear();
                else AppendLog("[エラー] タグの作成に失敗しました（同名タグが既に存在する可能性があります）。");
            }

            if (alsoPush)
            {
                AppendLog(string.Format("> GitHub へ送信中 (git push -u origin {0} --follow-tags)...", _branch));
                var pushRes = await _cmd.GitAsync(dir, false, "push", "-u", "origin", _branch, "--follow-tags");
                if (pushRes.Success) AppendLog("完了: ローカル保存および GitHub への同期が完了しました！");
                else AppendLog("[エラー] GitHub への送信に失敗しました。認証・ネットワーク・リモート設定を確認してください。");
            }
            else
            {
                AppendLog("完了: ローカル Git へのセーブが完了しました！");
            }
        }

        private async Task ExecutePushOnlyAsync()
        {
            string dir = _targetDir;
            if (!_hasRemote || _branch == null)
            {
                AppendLog("[エラー] リモート未設定、またはブランチ上にいないため送信できません。");
                return;
            }

            AppendLog("========== GitHub 送信 (Push) ==========");
            var res = await _cmd.GitAsync(dir, false, "push", "-u", "origin", _branch, "--follow-tags");
            if (res.Success) AppendLog("完了: GitHub へのバックアップ送信が完了しました！");
            else AppendLog("[エラー] GitHub への送信に失敗しました。認証・ネットワーク・リモート設定を確認してください。");
        }

        private async Task ExecutePullAsync()
        {
            string dir = _targetDir;
            if (IsExcelOpenInBin(dir)) return;
            if (!_hasRemote || _branch == null)
            {
                AppendLog("[エラー] リモート未設定、またはブランチ上にいないため取得できません。");
                return;
            }

            if (!await BackupBeforeOverwriteAsync(dir)) return;

            AppendLog("========== GitHub から最新取得 (Pull) ==========");
            var pullRes = await _cmd.GitAsync(dir, false, "pull", "origin", _branch);
            if (!pullRes.Success)
            {
                AppendLog("[エラー] git pull に失敗しました。ローカルに未保存の変更があるか、リモート/認証に問題があります。");
                return;
            }

            if (await RunVbacAsync(dir, "combine"))
                AppendLog("完了: 最新コードを取得し、bin内のExcelブックへ反映しました！");
            else
                AppendLog("[エラー] Excelブックへの書き戻し（combine）に失敗しました。");
        }

        // ---------------------------------------------------------------------
        // 復元
        // ---------------------------------------------------------------------
        private async Task<bool> ReplaceSrcWithRevisionAsync(string dir, string rev)
        {
            var chk = await _cmd.GitAsync(dir, true, "cat-file", "-e", rev + ":src");
            if (!chk.Success)
            {
                AppendLog("[エラー] 指定の履歴には src/ フォルダが存在しません。復元を中止しました。");
                return false;
            }

            await _cmd.GitAsync(dir, true, "rm", "-r", "-f", "-q", "--ignore-unmatch", "--", "src");
            await _cmd.GitAsync(dir, true, "clean", "-f", "-d", "-q", "--", "src");

            var co = await _cmd.GitAsync(dir, false, "checkout", rev, "--", "src");
            if (!co.Success)
            {
                AppendLog("[エラー] src/ の復元に失敗しました。HEAD からの復旧を試みます...");
                await _cmd.GitAsync(dir, false, "checkout", "HEAD", "--", "src");
                AppendLog("  退避フォルダ (.git\\vbagit_backup) に元の状態が残っています。");
                return false;
            }
            return true;
        }

        private async Task ExecuteRestoreLatestAsync()
        {
            string dir = _targetDir;
            if (IsExcelOpenInBin(dir)) return;

            if (MessageBox.Show("直前に Save した状態へソースコードとExcelブックを巻き戻します。\n現在の未保存の編集内容は破棄されます\n（ただし実行前に .git\\vbagit_backup へ自動退避します）。\nよろしいですか？",
                "直前状態への復元確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            if (!await BackupBeforeOverwriteAsync(dir)) return;

            AppendLog("========== 直前の Save 状態へ復元 ==========");
            if (!await ReplaceSrcWithRevisionAsync(dir, "HEAD")) return;

            if (await RunVbacAsync(dir, "combine"))
                AppendLog("完了: 直前の Save 状態へ Excel ブックを復元しました！");
            else
                AppendLog("[エラー] Excelブックへの書き戻し（combine）に失敗しました。");
        }

        private HistoryItem GetSelectedHistory()
        {
            var item = lstHistory.SelectedItem as HistoryItem;
            if (item == null || !Validation.IsValidCommitHash(item.Hash)) return null;
            return item;
        }

        private async Task ExecuteRestoreSelectedAsync(bool fromGithub)
        {
            string dir = _targetDir;
            if (IsExcelOpenInBin(dir)) return;

            var item = GetSelectedHistory();
            if (item == null)
            {
                MessageBox.Show("一覧から復元したい履歴を1つ選択してください。", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string commitHash = item.Hash;

            string sourceLabel = fromGithub ? "GitHub(origin)を取得したうえで、履歴" : "ローカルの履歴";
            if (MessageBox.Show(string.Format("{0} [{1}] の状態へ Excel ブックを復元します。\n現在の内容は .git\\vbagit_backup へ自動退避されます。\nよろしいですか？", sourceLabel, commitHash),
                "過去履歴からの復元確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            if (!await BackupBeforeOverwriteAsync(dir)) return;

            AppendLog(string.Format("========== 履歴 [{0}] へ復元開始 ==========", commitHash));
            if (fromGithub)
            {
                if (!_hasRemote)
                {
                    AppendLog("[エラー] リモートが未設定です。");
                    return;
                }
                var fetchRes = await _cmd.GitAsync(dir, false, "fetch", "origin", "--tags");
                if (!fetchRes.Success)
                {
                    AppendLog("[エラー] GitHub からの取得(fetch)に失敗しました。");
                    return;
                }
            }

            var verify = await _cmd.GitAsync(dir, true, "rev-parse", "--verify", "-q", commitHash + "^{commit}");
            if (!verify.Success)
            {
                AppendLog("[エラー] 指定されたコミットIDが見つかりません。");
                return;
            }

            if (!await ReplaceSrcWithRevisionAsync(dir, commitHash)) return;

            if (await RunVbacAsync(dir, "combine"))
            {
                AppendLog(string.Format("完了: [{0}] の状態を Excel ブックへ復元しました！", commitHash));
                AppendLog("※この状態を新しい正として確定させる場合は、確認後に「Save」または「Save ＆ GitHub送信」を押してください。");
            }
            else
            {
                AppendLog("[エラー] Excelブックへの書き戻し（combine）に失敗しました。");
            }
        }

        private async Task ExecuteTagSelectedAsync()
        {
            string dir = _targetDir;
            var item = GetSelectedHistory();
            if (item == null)
            {
                MessageBox.Show("タグを付けたい履歴を一覧から選択してください。", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string commitHash = item.Hash;

            string tagName = PromptInput("バージョンタグの付与", string.Format("選択履歴 [{0}] に付けるタグ名を入力してください\n（完成版や配布時の目印として使います。例: v1.0）:", commitHash), "v1.0");
            if (string.IsNullOrWhiteSpace(tagName)) return;
            tagName = tagName.Trim();

            if (!await IsValidTagNameAsync(dir, tagName))
            {
                MessageBox.Show("タグ名として使用できない文字が含まれています。\n（空白、~ ^ : ? * [ \\ などは使用できません）", "タグ名エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string tagMsg = PromptInput("タグのメモ", "タグの説明メモを入力してください:", "リリース版");
            tagMsg = string.IsNullOrWhiteSpace(tagMsg) ? tagName : Validation.SanitizeMessage(tagMsg);

            var res = await _cmd.GitAsync(dir, false, "tag", "-a", tagName, commitHash, "-m", tagMsg);
            if (res.Success)
                AppendLog(string.Format("完了: [{0}] にタグ '{1}' を付与しました。（GitHubへ送るには「GitHub送信のみ」を押してください）", commitHash, tagName));
            else
                AppendLog("[エラー] タグの付与に失敗しました（同名タグが既に存在する可能性があります）。");
        }

        // ---------------------------------------------------------------------
        // 新規作成 / クローン / 初期設定
        // ---------------------------------------------------------------------
        private async Task ExecuteInitLocalNewAsync()
        {
            string folderName = PromptInput("新規作成 (ローカル専用)", "作成する新しいプロジェクトのフォルダ名を入力してください\n（例: 売上集計ツール）:", "New_VBA_Project");
            if (string.IsNullOrWhiteSpace(folderName)) return;
            folderName = folderName.Trim();

            if (!Validation.IsValidFolderName(folderName))
            {
                MessageBox.Show("フォルダ名に使用できない文字が含まれています。\n（\\ / : * ? \" < > | など）", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string newProjectPath = Path.Combine(_targetDir, folderName);
            if (Directory.Exists(newProjectPath))
            {
                MessageBox.Show("同名のフォルダが既に存在します:\n" + newProjectPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            AppendLog(string.Format("========== ローカル専用プロジェクト作成 ({0}) ==========", folderName));
            Directory.CreateDirectory(newProjectPath);

            var res = await _cmd.GitAsync(newProjectPath, false, "init");
            if (!res.Success)
            {
                AppendLog("[エラー] git init に失敗しました。");
                return;
            }

            await _cmd.GitAsync(newProjectPath, true, "symbolic-ref", "HEAD", "refs/heads/main");

            ProjectFiles.EnsureDefaultGitignore(newProjectPath);
            AppendLog("・標準の .gitignore (bin/* , ~$*) と bin/.gitkeep を作成しました。");

            AppendLog("完了: ローカル専用のGit管理フォルダを作成しました！作業対象を切り替えます。");
            _targetDir = newProjectPath;
        }

        private async Task ExecuteCloneNewAsync()
        {
            string cloneUrl = PromptInput("新規クローン (1/2)", "GitHubの「Code」ボタンからコピーした\nリポジトリURLをそのまま貼り付けてください:\n（例: https://github.com/ユーザー名/リポジトリ名.git）", "");
            if (string.IsNullOrWhiteSpace(cloneUrl)) return;
            cloneUrl = cloneUrl.Trim();

            if (!Validation.IsSafeCloneUrl(cloneUrl))
            {
                MessageBox.Show("URLの形式が正しくありません。\n（空白・制御文字・先頭の「-」は使用できません）", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string defaultFolder = "";
            string lastPart = cloneUrl.TrimEnd('/').Split('/', ':').Last();
            if (lastPart.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
                lastPart = lastPart.Substring(0, lastPart.Length - 4);
            if (Validation.IsValidFolderName(lastPart)) defaultFolder = lastPart;

            string folderName = PromptInput("新規クローン (2/2)", "作成するローカルのフォルダ名を入力してください\n（そのままか、分かりやすい日本語名に変更可能です）:", defaultFolder);
            if (string.IsNullOrWhiteSpace(folderName)) return;
            folderName = folderName.Trim();

            if (!Validation.IsValidFolderName(folderName))
            {
                MessageBox.Show("フォルダ名に使用できない文字が含まれています。\n（\\ / : * ? \" < > | など）", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string newProjectPath = Path.Combine(_targetDir, folderName);
            if (Directory.Exists(newProjectPath))
            {
                MessageBox.Show("同名のフォルダが既に存在します:\n" + newProjectPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            AppendLog(string.Format("========== 新規クローン構築 ({0}) ==========", folderName));
            var res = await _cmd.GitAsync(_targetDir, false, "clone", "--", cloneUrl, folderName);
            if (!res.Success)
            {
                AppendLog("[エラー] クローンに失敗しました。URLと認証情報を確認してください。");
                return;
            }

            string binPath = Path.Combine(newProjectPath, "bin");
            if (!Directory.Exists(binPath)) Directory.CreateDirectory(binPath);

            AppendLog("完了: クローンが完了しました！作業対象を新しいフォルダへ切り替えます。");
            _targetDir = newProjectPath;
        }

        private async Task OpenInitProjectDialogAsync()
        {
            string dir = _targetDir;
            string gitignorePath = Path.Combine(dir, ".gitignore");
            if (!File.Exists(gitignorePath))
            {
                File.WriteAllText(gitignorePath, string.Join("\r\n", ProjectFiles.DefaultIgnoreLines) + "\r\n", new UTF8Encoding(false));
            }

            string newText;
            bool deleteTpl;
            string tplPath = Path.Combine(dir, "push_template.bat");

            using (Form dlg = new Form())
            {
                dlg.AutoScaleMode = AutoScaleMode.None;
                dlg.Font = this.Font;
                int fh = this.Font.Height;
                dlg.Text = "プロジェクト初期設定（.gitignore 編集 ＆ 監視解除）";
                dlg.ClientSize = new Size(Math.Max(600, fh * 34), Math.Max(460, fh * 26));
                dlg.MinimumSize = new Size(540, 400);
                dlg.StartPosition = FormStartPosition.CenterParent;

                var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var lbl = new Label { Text = "現在の .gitignore 内容（1行ずつ自由に編集・追加・削除できます）:", AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
                var txtIgnore = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 10f),
                    Text = File.ReadAllText(gitignorePath, Encoding.UTF8)
                };

                var btnAddPreset = new Button
                {
                    Text = "標準除外セットを追加 (*.bat / *.txt / vbac.wsf)",
                    AutoSize = true,
                    Padding = new Padding(8, 3, 8, 3),
                    BackColor = Color.White,
                    Margin = new Padding(0, 8, 0, 4)
                };
                btnAddPreset.Click += (s, e) =>
                {
                    var lines = txtIgnore.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Select(l => l.Trim()).ToList();
                    foreach (string item in new[] { "*.bat", "*.txt", "vbac.wsf" })
                    {
                        if (!lines.Contains(item)) lines.Add(item);
                    }
                    txtIgnore.Text = string.Join("\r\n", lines.Where(l => !string.IsNullOrEmpty(l)).ToArray()) + "\r\n";
                };

                bool tplExists = File.Exists(tplPath);
                var chkDelPushTpl = new CheckBox
                {
                    Text = "フォルダ内に push_template.bat が残っていたら削除する",
                    AutoSize = true,
                    Checked = tplExists,
                    Enabled = tplExists,
                    Margin = new Padding(2, 4, 0, 8)
                };

                var pnlBtns = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = true,
                    Margin = new Padding(0)
                };
                var btnOk = new Button
                {
                    Text = "確定してGit(ローカル/GitHub)から除外反映",
                    AutoSize = true,
                    Padding = new Padding(10, 5, 10, 5),
                    BackColor = BgBlue,
                    ForeColor = FgBlue,
                    Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold),
                    DialogResult = DialogResult.OK
                };
                var btnCancel = new Button
                {
                    Text = "キャンセル",
                    AutoSize = true,
                    Padding = new Padding(10, 5, 10, 5),
                    BackColor = Color.White,
                    DialogResult = DialogResult.Cancel,
                    Margin = new Padding(0, 0, 8, 0)
                };
                pnlBtns.Controls.Add(btnOk);
                pnlBtns.Controls.Add(btnCancel);

                tlp.Controls.Add(lbl, 0, 0);
                tlp.Controls.Add(txtIgnore, 0, 1);
                tlp.Controls.Add(btnAddPreset, 0, 2);
                tlp.Controls.Add(chkDelPushTpl, 0, 3);
                tlp.Controls.Add(pnlBtns, 0, 4);

                dlg.Controls.Add(tlp);
                dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                newText = txtIgnore.Text;
                deleteTpl = chkDelPushTpl.Checked;
            }

            AppendLog("========== 初期設定・監視解除を実行 ==========");
            File.WriteAllText(gitignorePath, newText.TrimEnd() + "\r\n", new UTF8Encoding(false));

            if (deleteTpl && File.Exists(tplPath))
            {
                File.Delete(tplPath);
                AppendLog("・不要な push_template.bat を削除しました。");
            }

            var lsRes = await _cmd.GitAsync(dir, true, "ls-files", "-i", "-c", "--exclude-standard");
            if (lsRes.Success)
            {
                var trackedIgnored = lsRes.Output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string file in trackedIgnored)
                {
                    string f = file.Trim();
                    if (f.Length == 0) continue;
                    await _cmd.GitAsync(dir, true, "rm", "--cached", "-f", "--", f);
                    AppendLog("・監視解除しました: " + f);
                }
            }

            await _cmd.GitAsync(dir, false, "add", ".gitignore");

            bool committed = false;
            if (await HasStagedChangesAsync(dir))
            {
                var cm = await _cmd.GitAsync(dir, false, "commit", "-m", "プロジェクト用.gitignore設定および不要ファイルの管理除外");
                committed = cm.Success;
                if (!committed) AppendLog("[エラー] コミットに失敗しました。user.name / user.email の設定を確認してください。");
            }
            else
            {
                AppendLog("[案内] 変更がないため、コミットはスキップされました。");
            }

            var br = await _cmd.GitAsync(dir, true, "symbolic-ref", "--short", "-q", "HEAD");
            string branch = br.Success ? br.Output.Trim() : "";
            var rm = await _cmd.GitAsync(dir, true, "remote", "get-url", "origin");
            bool hasRemote = rm.Success && !string.IsNullOrWhiteSpace(rm.Output);

            if (committed && hasRemote && branch.Length > 0)
            {
                var push = await _cmd.GitAsync(dir, false, "push", "-u", "origin", branch);
                if (push.Success) AppendLog("完了: .gitignore の更新と GitHub からの管理除外が完了しました！");
                else AppendLog("[エラー] GitHub への送信に失敗しました。後で「GitHub送信のみ」を実行してください。");
            }
            else
            {
                AppendLog("完了: .gitignore を保存し、ローカル履歴を更新しました。");
            }
        }

        private async Task RefreshHistoryListAsync()
        {
            lstHistory.Items.Clear();
            var res = await _cmd.GitAsync(_targetDir, true, "log", "-n", "30", "--date=short", "--pretty=format:%h%x1f%ad%x1f%d%x1f%s");
            if (!res.Success || string.IsNullOrWhiteSpace(res.Output)) return;

            var lines = res.Output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                string[] p = line.Split('\u001f');
                if (p.Length < 4 || !Validation.IsValidCommitHash(p[0])) continue;
                string deco = p[2].Trim();
                string display = p[0] + "  " + p[1] + "  " + (deco.Length > 0 ? deco + " " : "") + p[3];
                lstHistory.Items.Add(new HistoryItem { Hash = p[0], Display = display });
            }
        }

        // ---------------------------------------------------------------------
        // 右クリックメニュー (⑤ 常に塗りつぶしなし・黒文字・標準フォントを維持)
        // ---------------------------------------------------------------------
        private void UpdateRegButtonText()
        {
            if (ContextMenuRegistrar.IsRegistered())
            {
                btnToggleReg.Text = "右クリックメニュー登録済み (解除する)";
            }
            else
            {
                btnToggleReg.Text = "右クリックメニューに登録する";
            }
            btnToggleReg.BackColor = BgWhite;
            btnToggleReg.ForeColor = FgBlack;
            ApplyResponsiveLayout();
        }

        private void ToggleContextMenuRegistration()
        {
            if (_busy) return;
            try
            {
                if (ContextMenuRegistrar.IsRegistered())
                {
                    ContextMenuRegistrar.Unregister();
                    AppendLog("右クリックメニューから「VBA Git 管理ツール」を解除しました。");
                }
                else
                {
                    ContextMenuRegistrar.Register(Application.ExecutablePath);
                    AppendLog("右クリックメニューに「VBA Git 管理ツールを開く」を登録しました！（管理者権限不要・ユーザー領域）");
                }
                UpdateRegButtonText();
            }
            catch (Exception ex)
            {
                MessageBox.Show("レジストリ操作中にエラーが発生しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---------------------------------------------------------------------
        // ヘルパー
        // ---------------------------------------------------------------------
        private void AppendLog(string text)
        {
            if (txtLog == null || txtLog.IsDisposed) return;
            if (txtLog.InvokeRequired)
            {
                try { txtLog.BeginInvoke(new Action<string>(AppendLog), text); } catch { }
                return;
            }
            txtLog.AppendText(text + Environment.NewLine);
        }

        private string PromptInput(string title, string prompt, string defaultVal)
        {
            using (Form f = new Form())
            {
                f.AutoScaleMode = AutoScaleMode.None;
                f.Font = this.Font;
                int fh = this.Font.Height;
                f.Text = title;
                f.ClientSize = new Size(Math.Max(540, fh * 30), Math.Max(195, fh * 11));
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;

                var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(14) };
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var lbl = new Label { Text = prompt, Dock = DockStyle.Fill, AutoSize = true };
                var txt = new TextBox { Text = defaultVal, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 10) };

                var pnlBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Margin = new Padding(0) };
                var cancel = new Button { Text = "キャンセル", AutoSize = true, Padding = new Padding(10, 3, 10, 3), DialogResult = DialogResult.Cancel };
                var ok = new Button { Text = "OK", AutoSize = true, Padding = new Padding(16, 3, 16, 3), DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 8, 0) };
                pnlBtns.Controls.Add(cancel);
                pnlBtns.Controls.Add(ok);

                tlp.Controls.Add(lbl, 0, 0);
                tlp.Controls.Add(txt, 0, 1);
                tlp.Controls.Add(pnlBtns, 0, 2);

                f.Controls.Add(tlp);
                f.AcceptButton = ok;
                f.CancelButton = cancel;

                return f.ShowDialog(this) == DialogResult.OK ? txt.Text : null;
            }
        }
    }

    // =========================================================================
    // ヘルプ
    // =========================================================================
    public class HelpForm : Form
    {
        public HelpForm()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            this.Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular);
            int fh = this.Font.Height;

            this.Text = "VBA Git 管理ツール - 使い方ガイド";
            this.ClientSize = new Size(Math.Max(720, fh * 40), Math.Max(560, fh * 31));
            this.MinimumSize = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterParent;

            var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(10, 6) };

            tabs.TabPages.Add(CreateTab("1. 日常の開発サイクル", GetDailyGuideText()));
            tabs.TabPages.Add(CreateTab("2. 新規案件の立ち上げ", GetNewProjectGuideText()));
            tabs.TabPages.Add(CreateTab("3. やり直し・復元方法", GetRestoreGuideText()));
            tabs.TabPages.Add(CreateTab("4. ボタン機能一覧", GetReferenceGuideText()));

            var pnlBottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(10, 8, 10, 8)
            };
            var btnClose = new Button { Text = "閉じる", AutoSize = true, Padding = new Padding(14, 4, 14, 4), BackColor = Color.White };
            btnClose.Click += (s, e) => this.Close();
            pnlBottom.Controls.Add(btnClose);

            this.Controls.Add(tabs);
            this.Controls.Add(pnlBottom);
        }

        private TabPage CreateTab(string title, string content)
        {
            var page = new TabPage(title) { Padding = new Padding(8), BackColor = Color.White };
            var txt = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                Font = new Font("Yu Gothic UI", 10f),
                Text = content
            };
            page.Controls.Add(txt);
            return page;
        }

        private string GetDailyGuideText()
        {
            return
                "【日々のコーディング時の基本ステップ】\r\n\r\n" +
                "1. binフォルダ内のExcelでマクロを編集し、上書き保存（Ctrl + S）する\r\n\r\n" +
                "2. フォルダ内を右クリック ＞「VBA Git 管理ツールを開く」\r\n" +
                "   （※うっかり bin や src フォルダ内で起動しても、自動で親フォルダへ補正されます）\r\n\r\n" +
                "3.「コミットメモ」に修正内容を入力する\r\n" +
                "   （空欄のままでも自動で日時メモが入ります）\r\n\r\n" +
                "4. タイミングに合わせてボタンを押す\r\n" +
                "   ・PC内だけ（ローカル）に保存 → [Save (ローカル保存のみ)]\r\n" +
                "   ・GitHubにも送る（バックアップ） → [Save ＆ GitHub送信 (一括)]\r\n\r\n" +
                "※.gitignore に bin/ の除外設定が無い場合は、Save時に警告が表示されます。\r\n" +
                "※処理中はボタンが一時的に操作できなくなります（フォルダ変更・終了も不可）。\r\n\r\n" +
                "------------------------------------------------------------\r\n" +
                "【別のPCで作業を引き継ぐ時（GitHub利用時）】\r\n\r\n" +
                "1. 必ずExcelブックを閉じた状態にする\r\n" +
                "2. フォルダを右クリックしてツールを開き、[Pull (GitHub最新取得＋Excel反映)] を押す\r\n" +
                "3. GitHubから最新コードを取得し、binフォルダ内のExcelへ自動反映されます。\r\n" +
                "   （実行前に現在の状態が自動で退避されます）";
        }

        private string GetNewProjectGuideText()
        {
            return
                "【新しいVBA案件を始める時のステップ】\r\n\r\n" +
                "■ パターンA：手元のPC内だけで管理する（社外秘・単独開発など、最も一般的な使い方）\r\n" +
                "  1. 案件のフォルダを作りたい「親フォルダ」で右クリックして本ツールを開く\r\n" +
                "  2. [新規作成 (ローカル専用)] を押し、フォルダ名を入力する\r\n" +
                "  3. 新しいフォルダが作成され、自動的に切り替わります\r\n" +
                "     （標準の .gitignore と bin/.gitkeep も自動作成されます）\r\n" +
                "  4. 必要なら [初期設定 (.gitignore / 監視解除)] で除外設定を編集します\r\n" +
                "  5. 自動作成された bin フォルダにExcelファイルを入れ、[Save] を押せば初期登録完了です！\r\n\r\n" +
                "■ パターンB：GitHubを利用してバックアップや共有を行う場合\r\n" +
                "  1. GitHub上で新しい空のリポジトリを作成し、URLをコピーします\r\n" +
                "  2. 親フォルダで本ツールを開き、[新規クローン (GitHubから)] を押します\r\n" +
                "  3. コピーしたURLを貼り付けると、フォルダが作成され自動的に切り替わります\r\n" +
                "  4. 以降はパターンAの 4～5 と同じ手順です。";
        }

        private string GetRestoreGuideText()
        {
            return
                "【コードを壊してしまった・過去の状態に戻したい時】\r\n\r\n" +
                "※重要：復元系の操作を行う際は、必ず事前にExcelブックを閉じてください。\r\n" +
                "（開いたままだと安全装置が働き、警告ポップアップで自動停止します）\r\n\r\n" +
                "※復元・Pullの実行前に、現在の src と bin の内容が\r\n" +
                "  「.git\\vbagit_backup\\日時」へ自動退避されます（直近10回分を保持）。\r\n" +
                "  万一おかしくなった場合は、そこから手動で戻せます。\r\n\r\n" +
                "■ パターンA：昨日や先週など、特定の履歴・タグ時点に戻したい（最も使う復元）\r\n" +
                "  1.「過去の履歴・タグ一覧」から戻りたい行をクリックして選択します。\r\n" +
                "  2. [選択した履歴へ復元 (ローカル)] を押します（GitHub上の履歴を使う場合は [GitHub取得後]）。\r\n" +
                "  3. Excelを開いて動作を確認し、その状態を新しい最新版として確定させる場合は\r\n" +
                "     再度 [Save]（または GitHub送信）を押してください。\r\n\r\n" +
                "■ パターンB：さっきSaveした直前の状態にすぐ戻したい\r\n" +
                "  → [直前のSave状態へ戻す (ローカル)] を押します。\r\n" +
                "     最後に Save した時点のコードがExcelブックへ書き戻されます。\r\n\r\n" +
                "■ バージョンタグ（v1.0 など）を付けるタイミングと使い方\r\n" +
                "  日々の細かいSaveではなく、「v1.0完成！」「他部署への配布版」「大改修の直前」など、\r\n" +
                "  後からすぐに見つけ出したい重要な節目（マイルストーン）に「しおり」として使います。\r\n" +
                "  ・履歴一覧から行を選択して [選択履歴にタグ付与] を押し、タグ名を入力します。\r\n" +
                "  ・その後 [GitHub送信のみ (Push)] を押すとGitHub上にもタグが反映されます。";
        }

        private string GetReferenceGuideText()
        {
            return
                "【各ボタンの機能と安全装置について】\r\n\r\n" +
                "・Save (ローカル保存のみ)\r\n" +
                "  bin内のExcelからVBAコードを抽出し、手元のPC（ローカルGit）に履歴を記録します。\r\n\r\n" +
                "・Save ＆ GitHub送信 (一括) / GitHub送信のみ (Push)\r\n" +
                "  ローカルの履歴や後から付けたタグをGitHubへバックアップ送信します。\r\n\r\n" +
                "・選択した履歴へ復元 (ローカル) / 直前のSave状態へ戻す (ローカル)\r\n" +
                "  手元のPCに記録された履歴から、指定時点（または直前）のVBAコードをExcelへ書き戻します。\r\n\r\n" +
                "・Pull / 選択した履歴へ復元 (GitHub取得後)\r\n" +
                "  GitHub上の最新変更や過去履歴を取得してExcelへ反映します。\r\n\r\n" +
                "・選択履歴にタグ付与\r\n" +
                "  完成版や配布版など、重要な節目の履歴に「v1.0」等の目印（タグ）を付けます。\r\n\r\n" +
                "・親フォルダ自動補正（スマート判定）\r\n" +
                "  誤って bin や src フォルダ内で起動しても、1階層上の親フォルダを自動認識します。\r\n\r\n" +
                "・Excel使用中警告 ＆ 自動退避（安全装置）\r\n" +
                "  Excelが開いたままの復元を防止し、実行直前に .git\\vbagit_backup へ自動バックアップします。";
        }
    }
}