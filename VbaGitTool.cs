// VBA Git 管理ツール 修正版
// ビルド例 (C# 5 互換 / .NET Framework 4.x):
//   csc /target:winexe /out:VbaGitTool.exe Program.cs /r:System.Windows.Forms.dll /r:System.Drawing.dll
//
// 修正内容:
//  1. 復元時: 実行前に src/bin を .git\vbagit_backup\ へ自動退避 / 削除済みファイルが残らないよう src を置換
//  2. ブランチ名を動的取得（main 固定を廃止、detached HEAD / リモート無しを検知）
//  3. 引数エスケープ(Quote)・タグ名検証(check-ref-format)・フォルダ名/URL/ハッシュ検証
//  4. git 未インストール・実行例外・未処理例外のハンドリング
//  5. 処理中は操作ロック（フォルダ変更・終了も防止）
//  6. .gitignore 未設定時の Save 警告 / クラス分割 / DPI 対応

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

        /// <summary>CommandLineToArgvW 規則に従って 1 引数をクォートする。</summary>
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

        /// <summary>git を実行（引数は個別に渡す。自動でクォートされる）。</summary>
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
                // 非表示ウィンドウで資格情報入力待ちのまま固まるのを防ぐ
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
            if (url.StartsWith("-")) return false;          // オプション注入防止
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
    // 退避（バックアップ）/ .gitignore 補助
    // =========================================================================
    public static class ProjectFiles
    {
        public static readonly string[] DefaultIgnoreLines = { "bin/*", "!bin/.gitkeep", "~$*" };
        private const int KeepBackups = 10;

        /// <summary>
        /// src と bin を .git\vbagit_backup\yyyyMMdd_HHmmss へコピー。
        /// 戻り値: 退避先パス / 退避対象なしなら "" / 失敗なら null
        /// </summary>
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

                // 古い退避を削除
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

        /// <summary>.gitignore に標準行を（足りない分だけ）追加し、bin/.gitkeep を作る。BOMなしUTF-8。</summary>
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
    // メインフォーム
    // =========================================================================
    public class MainForm : Form
    {
        private string _targetDir;
        private readonly string _appDir;
        private readonly string _masterVbacPath;
        private readonly CommandRunner _cmd;

        // 状態
        private bool _busy;
        private bool _gitAvailable = true;
        private string _branch;       // null = detached HEAD / 不明
        private bool _hasRemote;

        // UI
        private Label lblTargetDir;
        private Label lblGitInfo;
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
            _targetDir = initialDir;
            _appDir = AppDomain.CurrentDomain.BaseDirectory;
            _masterVbacPath = Path.Combine(_appDir, "vbac.wsf");
            _cmd = new CommandRunner(AppendLog);

            InitializeComponents();
            this.Shown += async (s, e) => await InitialLoadAsync();
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

        // ---------------------------------------------------------------------
        // UI 構築
        // ---------------------------------------------------------------------
        private void InitializeComponents()
        {
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;

            this.Text = "VBA Git 管理ツール";
            this.Size = new Size(820, 740);
            this.MinimumSize = new Size(780, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 68, BackColor = Color.FromArgb(240, 244, 248), Padding = new Padding(10) };
            lblTargetDir = new Label { Location = new Point(12, 10), Size = new Size(530, 22), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), AutoEllipsis = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            lblGitInfo = new Label { Location = new Point(12, 36), Size = new Size(530, 22), ForeColor = Color.FromArgb(40, 80, 140), AutoEllipsis = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

            btnChangeDir = new Button { Text = "フォルダ変更...", Location = new Point(550, 16), Size = new Size(112, 34), BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnChangeDir.Click += (s, e) => ChangeTargetDirectory();

            btnHelp = new Button { Text = "ヘルプ / 使い方", Location = new Point(668, 16), Size = new Size(122, 34), BackColor = Color.FromArgb(230, 240, 255), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnHelp.Click += (s, e) => ShowHelpWindow();

            pnlTop.Controls.AddRange(new Control[] { lblTargetDir, lblGitInfo, btnChangeDir, btnHelp });

            // Group 1
            grpSave = new GroupBox { Text = "1. 日常のセーブ・GitHub同期", Location = new Point(12, 76), Size = new Size(780, 105), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            Label lblMsg = new Label { Text = "コミットメモ:", Location = new Point(14, 28), Size = new Size(85, 22) };
            txtCommitMsg = new TextBox { Location = new Point(102, 25), Size = new Size(440, 25), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            Label lblTag = new Label { Text = "タグ(任意):", Location = new Point(555, 28), Size = new Size(72, 22), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            txtTagName = new TextBox { Location = new Point(630, 25), Size = new Size(135, 25), Anchor = AnchorStyles.Top | AnchorStyles.Right };

            btnSave = new Button { Text = "Save (ローカル保存のみ)", Location = new Point(14, 60), Size = new Size(210, 34), BackColor = Color.FromArgb(230, 242, 255) };
            btnSave.Click += async (s, e) => await RunOperationAsync(() => ExecuteSaveAsync(false));

            btnSaveAndPush = new Button { Text = "Save ＆ GitHub送信 (一括)", Location = new Point(236, 60), Size = new Size(260, 34), BackColor = Color.FromArgb(215, 245, 225), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold) };
            btnSaveAndPush.Click += async (s, e) => await RunOperationAsync(() => ExecuteSaveAsync(true));

            btnPushOnly = new Button { Text = "GitHub送信のみ (Push)", Location = new Point(508, 60), Size = new Size(257, 34), BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnPushOnly.Click += async (s, e) => await RunOperationAsync(() => ExecutePushOnlyAsync());

            grpSave.Controls.AddRange(new Control[] { lblMsg, txtCommitMsg, lblTag, txtTagName, btnSave, btnSaveAndPush, btnPushOnly });

            // Group 2
            grpRestore = new GroupBox { Text = "2. 取得・復元（やり直し） ※Excelを閉じてから実行", Location = new Point(12, 188), Size = new Size(780, 225), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            btnPull = new Button { Text = "Pull (GitHubから最新取得 ＋ Excel反映)", Location = new Point(14, 24), Size = new Size(370, 34), BackColor = Color.FromArgb(255, 248, 220) };
            btnPull.Click += async (s, e) => await RunOperationAsync(() => ExecutePullAsync());

            btnRestoreLatest = new Button { Text = "直前のSave状態へ戻す (Latest)", Location = new Point(396, 24), Size = new Size(369, 34), BackColor = Color.FromArgb(255, 235, 235), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnRestoreLatest.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreLatestAsync());

            Label lblHist = new Label { Text = "過去の履歴・タグ一覧（選択して復元 or タグ付け）:", Location = new Point(14, 66), Size = new Size(400, 20) };
            Button btnRefreshHist = new Button { Text = "履歴再読込", Location = new Point(675, 62), Size = new Size(90, 24), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnRefreshHist.Click += async (s, e) => await RunOperationAsync(() => RefreshHistoryListAsync());

            lstHistory = new ListBox { Location = new Point(14, 88), Size = new Size(751, 88), Font = new Font("Consolas", 9.5f), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, HorizontalScrollbar = true };

            btnRestoreLocal = new Button { Text = "選択した履歴へ復元 (ローカル)", Location = new Point(14, 182), Size = new Size(250, 32), BackColor = Color.White };
            btnRestoreLocal.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreSelectedAsync(false));

            btnRestoreGithub = new Button { Text = "選択した履歴へ復元 (GitHubを取得後)", Location = new Point(274, 182), Size = new Size(270, 32), BackColor = Color.White };
            btnRestoreGithub.Click += async (s, e) => await RunOperationAsync(() => ExecuteRestoreSelectedAsync(true));

            btnAddTag = new Button { Text = "選択履歴にタグ付与", Location = new Point(554, 182), Size = new Size(211, 32), BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnAddTag.Click += async (s, e) => await RunOperationAsync(() => ExecuteTagSelectedAsync());

            Button[] restoreButtons = { btnPull, btnRestoreLatest, btnRefreshHist, btnRestoreLocal, btnRestoreGithub, btnAddTag };
            grpRestore.Controls.AddRange(new Control[] { lblHist, lstHistory });
            grpRestore.Controls.AddRange(restoreButtons);

            // Group 3
            grpSetup = new GroupBox { Text = "3. 新規構築・初期設定・メンテナンス", Location = new Point(12, 420), Size = new Size(780, 68), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

            btnInitLocalNew = new Button { Text = "新規作成 (ローカル専用)", Location = new Point(14, 24), Size = new Size(245, 34), BackColor = Color.White };
            btnInitLocalNew.Click += async (s, e) => await RunOperationAsync(() => ExecuteInitLocalNewAsync());

            btnCloneNew = new Button { Text = "新規クローン (GitHubから)", Location = new Point(269, 24), Size = new Size(245, 34), BackColor = Color.White };
            btnCloneNew.Click += async (s, e) => await RunOperationAsync(() => ExecuteCloneNewAsync());

            btnInitProject = new Button { Text = "初期設定 (.gitignore / 監視解除)", Location = new Point(524, 24), Size = new Size(241, 34), BackColor = Color.FromArgb(245, 250, 255), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnInitProject.Click += async (s, e) => await RunOperationAsync(() => OpenInitProjectDialogAsync());

            grpSetup.Controls.AddRange(new Control[] { btnInitLocalNew, btnCloneNew, btnInitProject });

            // Log
            Label lblLog = new Label { Text = "実行ログ:", Location = new Point(14, 496), Size = new Size(100, 20) };
            btnToggleReg = new Button { Text = "右クリックメニュー登録 / 解除", Location = new Point(570, 492), Size = new Size(222, 26), BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnToggleReg.Click += (s, e) => ToggleContextMenuRegistration();

            txtLog = new TextBox
            {
                Location = new Point(12, 520),
                Size = new Size(780, 170),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(220, 235, 220),
                Font = new Font("Consolas", 9.5f),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            this.Controls.AddRange(new Control[] { pnlTop, grpSave, grpRestore, grpSetup, lblLog, btnToggleReg, txtLog });
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
        /// <summary>
        /// すべての操作はここを通す。処理中は操作ロック、例外は握ってログへ、
        /// 終了時は必ず状態を再読込してUIを復帰させる。
        /// </summary>
        private async Task RunOperationAsync(Func<Task> body)
        {
            if (_busy) return;
            SetBusy(true);

            // C# 5 では catch / finally の中で await できないため、例外は変数に退避する
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

            // symbolic-ref はコミット未作成のブランチでも動く。detached HEAD では失敗する。
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
            btnInitProject.Enabled = isGitRepo;

            btnSaveAndPush.Enabled = isGitRepo && canSync;
            btnPushOnly.Enabled = isGitRepo && canSync;
            btnPull.Enabled = isGitRepo && canSync;
            btnRestoreGithub.Enabled = isGitRepo && hasRemote;
        }

        private void ChangeTargetDirectory()
        {
            if (_busy) return;
            using (var dlg = new FolderBrowserDialog { Description = "操作対象のフォルダを選択してください", SelectedPath = _targetDir })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _targetDir = dlg.SelectedPath;
                    txtLog.Clear();
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

        /// <summary>上書き系の処理の前に、src と bin を退避する。続行してよければ true。</summary>
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
            // 0 = 差分なし, 1 = 差分あり, それ以外 = エラー（安全側で「あり」と扱う）
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
                "[はい]　標準の除外設定（bin/* , ~$* ）を自動追加して続行\n" +
                "[いいえ]　除外設定なしで続行\n" +
                "[キャンセル]　Saveを中止",
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
        /// <summary>
        /// src/ を指定リビジョンの内容に完全置換する。
        /// （単なる checkout だと、その時点で存在しないファイルが src/ に残るため、いったん削除してから復元する）
        /// </summary>
        private async Task<bool> ReplaceSrcWithRevisionAsync(string dir, string rev)
        {
            // そのリビジョンに src/ があるか事前確認（無いのに消してしまう事故を防ぐ）
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

            string tagName = PromptInput("バージョンタグの付与", string.Format("選択履歴 [{0}] に付けるタグ名を入力してください（例: v1.0）:", commitHash), "v1.0");
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

            // コミット前でも、古いgitでも動く方法でブランチ名を main に統一
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
                dlg.AutoScaleDimensions = new SizeF(96F, 96F);
                dlg.AutoScaleMode = AutoScaleMode.Dpi;
                dlg.Font = this.Font;
                dlg.Text = "プロジェクト初期設定（.gitignore 編集 ＆ 監視解除）";
                dlg.Size = new Size(520, 440);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;

                Label lbl = new Label { Text = "現在の .gitignore 内容（1行ずつ自由に編集・追加・削除できます）:", Location = new Point(14, 12), Size = new Size(480, 22) };
                TextBox txtIgnore = new TextBox
                {
                    Location = new Point(14, 38),
                    Size = new Size(476, 220),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 10f),
                    Text = File.ReadAllText(gitignorePath, Encoding.UTF8)
                };

                Button btnAddPreset = new Button { Text = "標準除外セットを追加 (*.bat / *.txt / vbac.wsf)", Location = new Point(14, 268), Size = new Size(320, 30) };
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
                CheckBox chkDelPushTpl = new CheckBox
                {
                    Text = "フォルダ内に push_template.bat が残っていたら削除する",
                    Location = new Point(16, 308),
                    Size = new Size(460, 24),
                    Checked = tplExists,
                    Enabled = tplExists
                };

                Button btnOk = new Button { Text = "確定してGit(ローカル/GitHub)から除外反映", Location = new Point(220, 348), Size = new Size(270, 38), BackColor = Color.FromArgb(215, 245, 225), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), DialogResult = DialogResult.OK };
                Button btnCancel = new Button { Text = "キャンセル", Location = new Point(110, 348), Size = new Size(100, 38), DialogResult = DialogResult.Cancel };

                dlg.Controls.AddRange(new Control[] { lbl, txtIgnore, btnAddPreset, chkDelPushTpl, btnOk, btnCancel });
                dlg.AcceptButton = null;
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

            // 現在の状態を最新化して判定
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
        // 右クリックメニュー
        // ---------------------------------------------------------------------
        private void UpdateRegButtonText()
        {
            if (ContextMenuRegistrar.IsRegistered())
            {
                btnToggleReg.Text = "右クリックメニュー登録済み (解除する)";
                btnToggleReg.BackColor = Color.FromArgb(225, 245, 230);
            }
            else
            {
                btnToggleReg.Text = "右クリックメニューに登録する";
                btnToggleReg.BackColor = Color.FromArgb(255, 245, 225);
            }
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
                f.AutoScaleDimensions = new SizeF(96F, 96F);
                f.AutoScaleMode = AutoScaleMode.Dpi;
                f.Font = this.Font;
                f.Text = title;
                f.Size = new Size(540, 205);
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;

                Label lbl = new Label { Text = prompt, Location = new Point(14, 14), Size = new Size(500, 60) };
                TextBox txt = new TextBox { Text = defaultVal, Location = new Point(14, 82), Size = new Size(496, 25) };
                Button ok = new Button { Text = "OK", Location = new Point(340, 122), Size = new Size(80, 30), DialogResult = DialogResult.OK };
                Button cancel = new Button { Text = "キャンセル", Location = new Point(430, 122), Size = new Size(80, 30), DialogResult = DialogResult.Cancel };

                f.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
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
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Text = "VBA Git 管理ツール - 使い方ガイド";
            this.Size = new Size(680, 560);
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular);

            TabControl tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(10, 6) };

            tabs.TabPages.Add(CreateTab("1. 日常の開発サイクル", GetDailyGuideText()));
            tabs.TabPages.Add(CreateTab("2. 新規案件の立ち上げ", GetNewProjectGuideText()));
            tabs.TabPages.Add(CreateTab("3. やり直し・復元方法", GetRestoreGuideText()));
            tabs.TabPages.Add(CreateTab("4. ボタン機能一覧", GetReferenceGuideText()));

            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 48, BackColor = Color.FromArgb(245, 245, 245) };
            Button btnClose = new Button { Text = "閉じる", Location = new Point(550, 8), Size = new Size(100, 32), BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnClose.Click += (s, e) => this.Close();
            pnlBottom.Controls.Add(btnClose);

            this.Controls.Add(tabs);
            this.Controls.Add(pnlBottom);
        }

        private TabPage CreateTab(string title, string content)
        {
            TabPage page = new TabPage(title) { Padding = new Padding(8), BackColor = Color.White };
            TextBox txt = new TextBox
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
                "1. Excel側でマクロを編集し、上書き保存（Ctrl + S）する\r\n\r\n" +
                "2. フォルダ内を右クリック ＞「VBA Git 管理ツールを開く」\r\n\r\n" +
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
                "2. フォルダを右クリックしてツールを開き、[Pull (GitHubから最新取得 ＋ Excel反映)] を押す\r\n" +
                "3. GitHubから最新コードを取得し、binフォルダ内のExcelへ自動反映されます。\r\n" +
                "   （実行前に現在の状態が自動で退避されます）";
        }

        private string GetNewProjectGuideText()
        {
            return
                "【新しいVBA案件を始める時のステップ】\r\n\r\n" +
                "■ パターンA：手元のPC内だけで管理する（社外秘・外部サーバー不可の場合）\r\n" +
                "  1. 案件のフォルダを作りたい「親フォルダ」で右クリックして本ツールを開く\r\n" +
                "  2. [新規作成 (ローカル専用)] を押し、フォルダ名を入力する\r\n" +
                "  3. 新しいフォルダが作成され、自動的に切り替わります\r\n" +
                "     （標準の .gitignore と bin/.gitkeep も自動作成されます）\r\n" +
                "  4. 必要なら [初期設定 (.gitignore / 監視解除)] で除外設定を編集します\r\n" +
                "  5. binフォルダにExcelファイルを入れ、[Save] を押せば初期登録完了です！\r\n\r\n" +
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
                "■ パターンA：さっきSaveした直前の状態に戻したい\r\n" +
                "  → [直前のSave状態へ戻す (Latest)] を押します。\r\n" +
                "     最後に Save した時点のコードがExcelブックへ書き戻されます。\r\n\r\n" +
                "■ パターンB：昨日や先週など、特定の履歴・タグ時点に戻したい\r\n" +
                "  1. 画面中央の「過去の履歴・タグ一覧」から戻りたい行をクリックして選択します。\r\n" +
                "  2. [選択した履歴へ復元 (ローカル)] または [GitHubを取得後] を押します。\r\n" +
                "  3. Excelを開いて動作を確認し、その状態を新しい最新版として確定させる場合は\r\n" +
                "     再度 [Save]（または GitHub送信）を押してください。\r\n\r\n" +
                "■ バージョンタグ（v1.0 など）を後から付けたい時\r\n" +
                "  ・履歴一覧から行を選択して [選択履歴にタグ付与] を押し、タグ名を入力します。\r\n" +
                "  ・その後 [GitHub送信のみ (Push)] を押すとGitHub上にもタグが反映されます。";
        }

        private string GetReferenceGuideText()
        {
            return
                "【各ボタンの機能と安全装置について】\r\n\r\n" +
                "・Save (ローカル保存のみ)\r\n" +
                "  bin内のExcelからVBAコードを抽出し、手元のPC（ローカルGit）に履歴を記録します。\r\n\r\n" +
                "・Save ＆ GitHub送信 (一括)\r\n" +
                "  上記のSaveに加え、GitHubへバックアップ送信（Push）まで一気に行います。\r\n" +
                "  （リモート未設定・detached HEAD の場合は無効になります）\r\n\r\n" +
                "・GitHub送信のみ (Push)\r\n" +
                "  ローカルにSave済みの履歴や、後から付けたタグをGitHubへ送信します。\r\n\r\n" +
                "・新規作成 / 新規クローン / 初期設定 (.gitignore / 監視解除)\r\n" +
                "  空のプロジェクト作成、URLからの環境複製、除外ファイル設定を行います。\r\n\r\n" +
                "・Excel使用中警告（安全装置）\r\n" +
                "  Pullや復元時にbin内にロックファイル（~$*.xlsm）がある場合は処理を中断します。\r\n\r\n" +
                "・自動退避\r\n" +
                "  Pull / 復元の直前に src と bin を .git\\vbagit_backup へコピーします。";
        }
    }
}
