using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VbaGitTool
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string targetDir = Environment.CurrentDirectory;
            if (args != null && args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                string rawArg = args[0].Trim().Trim('"');
                if (Directory.Exists(rawArg))
                {
                    targetDir = rawArg;
                }
            }

            Application.Run(new MainForm(targetDir));
        }
    }

    public class MainForm : Form
    {
        private string _targetDir;
        private readonly string _appDir;
        private readonly string _exePath;
        private readonly string _masterVbacPath;

        // UI Controls
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
        private Button btnInitLocalNew; // NEW
        private Button btnCloneNew;
        private Button btnInitProject;
        private Button btnPushTemplate;
        private Button btnToggleReg;
        private Button btnChangeDir;
        private Button btnHelp;
        private TextBox txtLog;
        private HelpForm _helpForm;

        private const string RegKeyBg = @"Software\Classes\Directory\Background\shell\VbaGitTool";
        private const string RegKeyDir = @"Software\Classes\Directory\shell\VbaGitTool";

        public MainForm(string initialDir)
        {
            _targetDir = initialDir;
            _appDir = AppDomain.CurrentDomain.BaseDirectory;
            _exePath = Process.GetCurrentProcess().MainModule.FileName;
            _masterVbacPath = Path.Combine(_appDir, "vbac.wsf");

            InitializeComponents();
            LoadDirectoryContext();
        }

        private void InitializeComponents()
        {
            this.Text = "VBA Git 管理ツール";
            this.Size = new Size(820, 780); // 高さを拡張
            this.MinimumSize = new Size(780, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular);

            // Top Info Panel
            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 68, BackColor = Color.FromArgb(240, 244, 248), Padding = new Padding(10) };
            lblTargetDir = new Label { Location = new Point(12, 10), Size = new Size(530, 22), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), AutoEllipsis = true };
            lblGitInfo = new Label { Location = new Point(12, 36), Size = new Size(530, 22), ForeColor = Color.FromArgb(40, 80, 140) };

            btnChangeDir = new Button { Text = "フォルダ変更...", Location = new Point(550, 16), Size = new Size(112, 34), BackColor = Color.White };
            btnChangeDir.Click += (s, e) => ChangeTargetDirectory();

            btnHelp = new Button { Text = "ヘルプ / 使い方", Location = new Point(668, 16), Size = new Size(122, 34), BackColor = Color.FromArgb(230, 240, 255), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold) };
            btnHelp.Click += (s, e) => ShowHelpWindow();

            pnlTop.Controls.AddRange(new Control[] { lblTargetDir, lblGitInfo, btnChangeDir, btnHelp });

            // Group 1: Save & Sync
            GroupBox grpSave = new GroupBox { Text = "1. 日常のセーブ・GitHub同期", Location = new Point(12, 76), Size = new Size(780, 105) };
            Label lblMsg = new Label { Text = "コミットメモ:", Location = new Point(14, 28), Size = new Size(85, 22) };
            txtCommitMsg = new TextBox { Location = new Point(102, 25), Size = new Size(440, 25) };
            Label lblTag = new Label { Text = "タグ(任意):", Location = new Point(555, 28), Size = new Size(72, 22) };
            txtTagName = new TextBox { Location = new Point(630, 25), Size = new Size(135, 25) };

            btnSave = new Button { Text = "Save (ローカル保存のみ)", Location = new Point(14, 60), Size = new Size(210, 34), BackColor = Color.FromArgb(230, 242, 255) };
            btnSave.Click += async (s, e) => await ExecuteSaveAsync(false);

            btnSaveAndPush = new Button { Text = "Save ＆ GitHub送信 (一括)", Location = new Point(236, 60), Size = new Size(260, 34), BackColor = Color.FromArgb(215, 245, 225), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold) };
            btnSaveAndPush.Click += async (s, e) => await ExecuteSaveAsync(true);

            btnPushOnly = new Button { Text = "GitHub送信のみ (Push)", Location = new Point(508, 60), Size = new Size(257, 34), BackColor = Color.White };
            btnPushOnly.Click += async (s, e) => await ExecutePushOnlyAsync();

            grpSave.Controls.AddRange(new Control[] { lblMsg, txtCommitMsg, lblTag, txtTagName, btnSave, btnSaveAndPush, btnPushOnly });

            // Group 2: Pull & Restore
            GroupBox grpRestore = new GroupBox { Text = "2. 取得・復元（やり直し） ※Excelを閉じてから実行", Location = new Point(12, 188), Size = new Size(780, 225) };
            btnPull = new Button { Text = "Pull (GitHubから最新取得 ＋ Excel反映)", Location = new Point(14, 24), Size = new Size(370, 34), BackColor = Color.FromArgb(255, 248, 220) };
            btnPull.Click += async (s, e) => await ExecutePullAsync();

            btnRestoreLatest = new Button { Text = "直前のSave状態へ戻す (Latest)", Location = new Point(396, 24), Size = new Size(369, 34), BackColor = Color.FromArgb(255, 235, 235) };
            btnRestoreLatest.Click += async (s, e) => await ExecuteRestoreLatestAsync();

            Label lblHist = new Label { Text = "過去の履歴・タグ一覧（選択して復元 or タグ付け）:", Location = new Point(14, 66), Size = new Size(400, 20) };
            Button btnRefreshHist = new Button { Text = "履歴再読込", Location = new Point(675, 62), Size = new Size(90, 24) };
            btnRefreshHist.Click += async (s, e) => await RefreshHistoryListAsync();

            lstHistory = new ListBox { Location = new Point(14, 88), Size = new Size(751, 88), Font = new Font("Consolas", 9.5f) };

            btnRestoreLocal = new Button { Text = "選択した履歴へ復元 (ローカル)", Location = new Point(14, 182), Size = new Size(250, 32), BackColor = Color.White };
            btnRestoreLocal.Click += async (s, e) => await ExecuteRestoreSelectedAsync(false);

            btnRestoreGithub = new Button { Text = "選択した履歴へ復元 (GitHubから取得)", Location = new Point(274, 182), Size = new Size(270, 32), BackColor = Color.White };
            btnRestoreGithub.Click += async (s, e) => await ExecuteRestoreSelectedAsync(true);

            btnAddTag = new Button { Text = "選択履歴にタグ付与", Location = new Point(554, 182), Size = new Size(211, 32), BackColor = Color.White };
            btnAddTag.Click += async (s, e) => await ExecuteTagSelectedAsync();

            grpRestore.Controls.AddRange(new Control[] { btnPull, btnRestoreLatest, lblHist, btnRefreshHist, lstHistory, btnRestoreLocal, btnRestoreGithub, btnAddTag });

            // Group 3: Setup & Maintenance (2段に拡張)
            GroupBox grpSetup = new GroupBox { Text = "3. 新規構築・初期設定・メンテナンス", Location = new Point(12, 420), Size = new Size(780, 105) };
            
            btnInitLocalNew = new Button { Text = "新規作成 (ローカル専用)", Location = new Point(14, 24), Size = new Size(370, 34), BackColor = Color.White };
            btnInitLocalNew.Click += async (s, e) => await ExecuteInitLocalNewAsync();

            btnCloneNew = new Button { Text = "新規クローン (GitHubから)", Location = new Point(396, 24), Size = new Size(369, 34), BackColor = Color.White };
            btnCloneNew.Click += async (s, e) => await ExecuteCloneNewAsync();

            btnInitProject = new Button { Text = "初期設定 (.gitignore / 監視解除)", Location = new Point(14, 62), Size = new Size(370, 34), BackColor = Color.White };
            btnInitProject.Click += async (s, e) => await OpenInitProjectDialogAsync();

            btnPushTemplate = new Button { Text = "テンプレート更新 (VBA_Template専用)", Location = new Point(396, 62), Size = new Size(369, 34), BackColor = Color.FromArgb(245, 235, 255) };
            btnPushTemplate.Click += async (s, e) => await ExecutePushTemplateAsync();

            grpSetup.Controls.AddRange(new Control[] { btnInitLocalNew, btnCloneNew, btnInitProject, btnPushTemplate });

            // Log Area (Y座標を全体的に下へシフト)
            Label lblLog = new Label { Text = "実行ログ:", Location = new Point(14, 536), Size = new Size(100, 20) };
            btnToggleReg = new Button { Text = "右クリックメニュー登録 / 解除", Location = new Point(570, 532), Size = new Size(222, 26), BackColor = Color.White };
            btnToggleReg.Click += (s, e) => ToggleContextMenuRegistration();

            txtLog = new TextBox
            {
                Location = new Point(12, 560),
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

        private async void LoadDirectoryContext()
        {
            lblTargetDir.Text = "対象フォルダ: " + _targetDir;
            UpdateRegButtonText();

            if (!File.Exists(_masterVbacPath))
            {
                AppendLog("[警告] アプリと同じフォルダに vbac.wsf が見つかりません: " + _masterVbacPath);
            }

            bool isGitRepo = Directory.Exists(Path.Combine(_targetDir, ".git"));
            if (!isGitRepo)
            {
                lblGitInfo.Text = "状態: Git未初期化フォルダ（親フォルダとして「新規作成」「クローン」が可能です）";
                SetProjectButtonsEnabled(false, false);
                lstHistory.Items.Clear();
                return;
            }

            string branch = (await RunCommandAsync("git", "rev-parse --abbrev-ref HEAD", _targetDir, true)).Output.Trim();
            string remoteUrl = (await RunCommandAsync("git", "remote get-url origin", _targetDir, true)).Output.Trim();
            if (string.IsNullOrEmpty(remoteUrl)) remoteUrl = "(リモート未設定・ローカル専用)";

            lblGitInfo.Text = string.Format("ブランチ: {0}   |   接続先(origin): {1}", branch, remoteUrl);

            bool isTemplateRepo = remoteUrl.IndexOf("VBA_Template", StringComparison.OrdinalIgnoreCase) >= 0;
            SetProjectButtonsEnabled(true, isTemplateRepo);

            await RefreshHistoryListAsync();
        }

        private void SetProjectButtonsEnabled(bool isGitRepo, bool isTemplateRepo)
        {
            // プロジェクト内部で誤って新規構築しないよう、Git未初期化フォルダでのみ新規作成ボタンを有効化
            btnInitLocalNew.Enabled = !isGitRepo;
            btnCloneNew.Enabled = !isGitRepo;
            
            btnSave.Enabled = isGitRepo && !isTemplateRepo;
            btnSaveAndPush.Enabled = isGitRepo && !isTemplateRepo;
            btnPushOnly.Enabled = isGitRepo;
            btnPull.Enabled = isGitRepo && !isTemplateRepo;
            btnRestoreLatest.Enabled = isGitRepo && !isTemplateRepo;
            btnRestoreLocal.Enabled = isGitRepo && !isTemplateRepo;
            btnRestoreGithub.Enabled = isGitRepo && !isTemplateRepo;
            btnAddTag.Enabled = isGitRepo;
            btnInitProject.Enabled = isGitRepo && !isTemplateRepo;
            btnPushTemplate.Enabled = isGitRepo && isTemplateRepo;
        }

        private void ChangeTargetDirectory()
        {
            using (var dlg = new FolderBrowserDialog { Description = "操作対象のフォルダを選択してください", SelectedPath = _targetDir })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _targetDir = dlg.SelectedPath;
                    txtLog.Clear();
                    LoadDirectoryContext();
                }
            }
        }

        // =====================================================================
        // Safety Checks
        // =====================================================================
        private bool IsExcelOpenInBin()
        {
            string binDir = Path.Combine(_targetDir, "bin");
            if (!Directory.Exists(binDir)) return false;

            var lockFiles = Directory.GetFiles(binDir, "~$*");
            if (lockFiles.Length > 0)
            {
                string fileNames = string.Join("\n", lockFiles.Select(Path.GetFileName));
                MessageBox.Show(
                    "binフォルダ内のExcelブックが開かれています！\n\n検出された一時ファイル:\n" + fileNames + "\n\nExcelを完全に閉じてから再実行してください。",
                    "Excel使用中警告（安全装置）",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return true;
            }
            return false;
        }

        private async Task<bool> RunVbacAsync(string mode)
        {
            if (!File.Exists(_masterVbacPath))
            {
                MessageBox.Show("アプリフォルダ内に vbac.wsf が存在しません。\n配置先: " + _masterVbacPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string binDir = Path.Combine(_targetDir, "bin");
            if (!Directory.Exists(binDir)) Directory.CreateDirectory(binDir);

            string tempVbac = Path.Combine(_targetDir, "vbac.wsf");
            bool copiedTemp = false;
            try
            {
                if (!File.Exists(tempVbac))
                {
                    File.Copy(_masterVbacPath, tempVbac, true);
                    copiedTemp = true;
                }

                AppendLog(string.Format("> cscript //nologo vbac.wsf {0}", mode));
                var res = await RunCommandAsync("cscript", "//nologo vbac.wsf " + mode, _targetDir, false);
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

        // =====================================================================
        // Core Operations
        // =====================================================================
        private async Task ExecuteSaveAsync(bool alsoPush)
        {
            SetAllButtonsEnabled(false);
            try
            {
                AppendLog("========== Save 処理開始 ==========");
                bool vbacOk = await RunVbacAsync("decombine");
                if (!vbacOk)
                {
                    AppendLog("[エラー] vbac.wsf decombine に失敗しました。");
                    return;
                }

                await RunCommandAsync("git", "add -A", _targetDir, false);

                string msg = txtCommitMsg.Text.Trim();
                if (string.IsNullOrEmpty(msg))
                {
                    msg = "作業セーブ (" + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ")";
                }

                var commitRes = await RunCommandAsync("git", string.Format("commit -m \"{0}\"", msg.Replace("\"", "'")), _targetDir, false);
                if (commitRes.ExitCode != 0)
                {
                    AppendLog("[案内] 前回セーブからコードの変更点がないため、新規コミットはスキップされました。");
                }
                else
                {
                    txtCommitMsg.Clear();
                }

                string tag = txtTagName.Text.Trim();
                if (!string.IsNullOrEmpty(tag))
                {
                    await RunCommandAsync("git", string.Format("tag -a \"{0}\" -m \"{1}\"", tag, msg.Replace("\"", "'")), _targetDir, false);
                    txtTagName.Clear();
                }

                if (alsoPush)
                {
                    AppendLog("> GitHub へ送信中 (git push origin main --follow-tags)...");
                    var pushRes = await RunCommandAsync("git", "push origin main --follow-tags", _targetDir, false);
                    if (pushRes.ExitCode == 0)
                        AppendLog("完了: ローカル保存および GitHub への同期が完了しました！");
                    else
                        AppendLog("[エラー] GitHub への送信に失敗しました。リモートが設定されていないか、ネットワークを確認してください。");
                }
                else
                {
                    AppendLog("完了: ローカル Git へのセーブが完了しました！");
                }

                await RefreshHistoryListAsync();
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task ExecutePushOnlyAsync()
        {
            SetAllButtonsEnabled(false);
            try
            {
                AppendLog("========== GitHub 送信 (Push) ==========");
                var res = await RunCommandAsync("git", "push origin main --follow-tags", _targetDir, false);
                if (res.ExitCode == 0)
                    AppendLog("完了: GitHub へのバックアップ送信が完了しました！");
                else
                    AppendLog("[エラー] GitHub への送信に失敗しました。リモート設定がない可能性があります。");
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task ExecutePullAsync()
        {
            if (IsExcelOpenInBin()) return;

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog("========== GitHub から最新取得 (Pull) ==========");
                var pullRes = await RunCommandAsync("git", "pull origin main", _targetDir, false);
                if (pullRes.ExitCode != 0)
                {
                    AppendLog("[エラー] git pull に失敗しました。ローカルに未保存の変更があるか、リモート設定がありません。");
                    return;
                }

                bool vbacOk = await RunVbacAsync("combine");
                if (vbacOk)
                    AppendLog("完了: 最新コードを取得し、bin内のExcelブックへ反映しました！");
                else
                    AppendLog("[エラー] Excelブックへの書き戻し（combine）に失敗しました。");

                await RefreshHistoryListAsync();
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task ExecuteRestoreLatestAsync()
        {
            if (IsExcelOpenInBin()) return;

            if (MessageBox.Show("直前に Save した状態へソースコードとExcelブックを巻き戻します。\n現在の未保存の編集内容は破棄されます。よろしいですか？",
                "直前状態への復元確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog("========== 直前の Save 状態へ復元 ==========");
                await RunCommandAsync("git", "checkout HEAD -- src/", _targetDir, false);
                bool vbacOk = await RunVbacAsync("combine");
                if (vbacOk)
                    AppendLog("完了: 直前の Save 状態へ Excel ブックを復元しました！");
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task ExecuteRestoreSelectedAsync(bool fromGithub)
        {
            if (IsExcelOpenInBin()) return;

            if (lstHistory.SelectedItem == null)
            {
                MessageBox.Show("一覧から復元したい履歴を1つ選択してください。", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string line = lstHistory.SelectedItem.ToString();
            string commitHash = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];

            string sourceLabel = fromGithub ? "GitHub(origin)の履歴" : "ローカルの履歴";
            if (MessageBox.Show(string.Format("{0} [{1}] の状態へ Excel ブックを復元します。\nよろしいですか？", sourceLabel, commitHash),
                "過去履歴からの復元確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog(string.Format("========== 履歴 [{0}] へ復元開始 ==========", commitHash));
                if (fromGithub)
                {
                    await RunCommandAsync("git", "fetch origin", _targetDir, false);
                }

                var coRes = await RunCommandAsync("git", string.Format("checkout {0} -- src/", commitHash), _targetDir, false);
                if (coRes.ExitCode != 0)
                {
                    AppendLog("[エラー] 指定されたコミットIDの取得に失敗しました。");
                    return;
                }

                bool vbacOk = await RunVbacAsync("combine");
                if (vbacOk)
                {
                    AppendLog(string.Format("完了: [{0}] の状態を Excel ブックへ復元しました！", commitHash));
                    AppendLog("※この状態を新しい正として確定させる場合は、確認後に「Save & GitHub送信」を押してください。");
                }
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task ExecuteTagSelectedAsync()
        {
            if (lstHistory.SelectedItem == null)
            {
                MessageBox.Show("タグを付けたい履歴を一覧から選択してください。", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string line = lstHistory.SelectedItem.ToString();
            string commitHash = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];

            string tagName = PromptInput("バージョンタグの付与", string.Format("選択履歴 [{0}] に付けるタグ名を入力してください（例: v1.0）:", commitHash), "v1.0");
            if (string.IsNullOrWhiteSpace(tagName)) return;

            string tagMsg = PromptInput("タグのメモ", "タグの説明メモを入力してください:", "リリース版");
            if (tagMsg == null) tagMsg = tagName;

            SetAllButtonsEnabled(false);
            try
            {
                var res = await RunCommandAsync("git", string.Format("tag -a \"{0}\" {1} -m \"{2}\"", tagName.Trim(), commitHash, tagMsg.Replace("\"", "'")), _targetDir, false);
                if (res.ExitCode == 0)
                {
                    AppendLog(string.Format("完了: [{0}] にタグ '{1}' を付与しました。（GitHubへ送るには「GitHub送信のみ」を押してください）", commitHash, tagName.Trim()));
                    await RefreshHistoryListAsync();
                }
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        // =====================================================================
        // NEW: ローカル専用での新規作成（git init）
        // =====================================================================
        private async Task ExecuteInitLocalNewAsync()
        {
            string folderName = PromptInput("新規作成 (ローカル専用)", "作成する新しいプロジェクトのフォルダ名を入力してください\n（例: 売上集計ツール）:", "New_VBA_Project");
            if (string.IsNullOrWhiteSpace(folderName)) return;
            folderName = folderName.Trim();

            string newProjectPath = Path.Combine(_targetDir, folderName);
            if (Directory.Exists(newProjectPath))
            {
                MessageBox.Show("同名のフォルダが既に存在します:\n" + newProjectPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog(string.Format("========== ローカル専用プロジェクト作成 ({0}) ==========", folderName));
                Directory.CreateDirectory(newProjectPath);

                var res = await RunCommandAsync("git", "init", newProjectPath, false);
                if (res.ExitCode != 0)
                {
                    AppendLog("[エラー] git init に失敗しました。");
                    return;
                }
                
                await RunCommandAsync("git", "branch -M main", newProjectPath, true);

                string binPath = Path.Combine(newProjectPath, "bin");
                Directory.CreateDirectory(binPath);

                AppendLog("完了: ローカル専用のGit管理フォルダを作成しました！作業対象を切り替えます。");
                _targetDir = newProjectPath;
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        // =====================================================================
        // URL貼り付け方式の新規クローン処理
        // =====================================================================
        private async Task ExecuteCloneNewAsync()
        {
            string cloneUrl = PromptInput("新規クローン (1/2)", "GitHubの「Code」ボタンからコピーした\nリポジトリURLをそのまま貼り付けてください:\n（例: https://github.com/ユーザー名/リポジトリ名.git）", "");
            if (string.IsNullOrWhiteSpace(cloneUrl)) return;
            cloneUrl = cloneUrl.Trim();

            string defaultFolder = "";
            if (cloneUrl.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
            {
                var parts = cloneUrl.Split('/');
                defaultFolder = parts.Last().Replace(".git", "");
            }

            string folderName = PromptInput("新規クローン (2/2)", "作成するローカルのフォルダ名を入力してください\n（そのままか、分かりやすい日本語名に変更可能です）:", defaultFolder);
            if (string.IsNullOrWhiteSpace(folderName)) return;
            folderName = folderName.Trim();

            string newProjectPath = Path.Combine(_targetDir, folderName);
            if (Directory.Exists(newProjectPath))
            {
                MessageBox.Show("同名のフォルダが既に存在します:\n" + newProjectPath, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog(string.Format("========== 新規クローン構築 ({0}) ==========", folderName));
                var res = await RunCommandAsync("git", string.Format("clone \"{0}\" \"{1}\"", cloneUrl, folderName), _targetDir, false);
                if (res.ExitCode != 0)
                {
                    AppendLog("[エラー] クローンに失敗しました。URLが正しいか確認してください。");
                    return;
                }

                string binPath = Path.Combine(newProjectPath, "bin");
                if (!Directory.Exists(binPath)) Directory.CreateDirectory(binPath);

                AppendLog("完了: クローンが完了しました！作業対象を新しいフォルダへ切り替えます。");
                _targetDir = newProjectPath;
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task OpenInitProjectDialogAsync()
        {
            string gitignorePath = Path.Combine(_targetDir, ".gitignore");
            if (!File.Exists(gitignorePath))
            {
                File.WriteAllText(gitignorePath, "bin/*\r\n!bin/.gitkeep\r\n~$*\r\n", Encoding.UTF8);
            }

            using (Form dlg = new Form())
            {
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
                    txtIgnore.Text = string.Join("\r\n", lines.Where(l => !string.IsNullOrEmpty(l))) + "\r\n";
                };

                CheckBox chkDelPushTpl = new CheckBox
                {
                    Text = "フォルダ内に push_template.bat が残っていたら削除する",
                    Location = new Point(16, 308),
                    Size = new Size(460, 24),
                    Checked = File.Exists(Path.Combine(_targetDir, "push_template.bat")),
                    Enabled = File.Exists(Path.Combine(_targetDir, "push_template.bat"))
                };

                Button btnOk = new Button { Text = "確定してGit(ローカル/GitHub)から除外反映", Location = new Point(220, 348), Size = new Size(270, 38), BackColor = Color.FromArgb(215, 245, 225), Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Bold), DialogResult = DialogResult.OK };
                Button btnCancel = new Button { Text = "キャンセル", Location = new Point(110, 348), Size = new Size(100, 38), DialogResult = DialogResult.Cancel };

                dlg.Controls.AddRange(new Control[] { lbl, txtIgnore, btnAddPreset, chkDelPushTpl, btnOk, btnCancel });

                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                SetAllButtonsEnabled(false);
                try
                {
                    AppendLog("========== 初期設定・監視解除を実行 ==========");
                    File.WriteAllText(gitignorePath, txtIgnore.Text.TrimEnd() + "\r\n", new UTF8Encoding(false));

                    if (chkDelPushTpl.Checked)
                    {
                        string tplBat = Path.Combine(_targetDir, "push_template.bat");
                        if (File.Exists(tplBat))
                        {
                            File.Delete(tplBat);
                            AppendLog("・不要な push_template.bat を削除しました。");
                        }
                    }

                    var lsRes = await RunCommandAsync("git", "ls-files -i -c --exclude-from=.gitignore", _targetDir, true);
                    var trackedIgnored = lsRes.Output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string file in trackedIgnored)
                    {
                        await RunCommandAsync("git", string.Format("rm --cached -f \"{0}\"", file.Trim()), _targetDir, true);
                        AppendLog("・監視解除しました: " + file.Trim());
                    }

                    await RunCommandAsync("git", "add .gitignore", _targetDir, false);
                    var cmRes = await RunCommandAsync("git", "commit -m \"プロジェクト用.gitignore設定および不要ファイルの管理除外\"", _targetDir, false);
                    
                    // リモートが設定されていればプッシュも試みる
                    string remoteUrl = (await RunCommandAsync("git", "remote get-url origin", _targetDir, true)).Output.Trim();
                    if (cmRes.ExitCode == 0 && !string.IsNullOrEmpty(remoteUrl))
                    {
                        await RunCommandAsync("git", "push origin main", _targetDir, false);
                        AppendLog("完了: .gitignore の更新と GitHub からの管理除外が完了しました！");
                    }
                    else
                    {
                        AppendLog("完了: .gitignore を保存し、ローカル履歴を更新しました。");
                    }
                }
                finally
                {
                    LoadDirectoryContext();
                }
            }
        }

        private async Task ExecutePushTemplateAsync()
        {
            string msg = txtCommitMsg.Text.Trim();
            if (string.IsNullOrEmpty(msg)) msg = "マスターテンプレート更新 (" + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ")";

            SetAllButtonsEnabled(false);
            try
            {
                AppendLog("========== マスターテンプレート (VBA_Template) 更新 ==========");
                await RunCommandAsync("git", "add -A", _targetDir, false);
                await RunCommandAsync("git", string.Format("commit -m \"{0}\"", msg.Replace("\"", "'")), _targetDir, false);
                var pushRes = await RunCommandAsync("git", "push -u origin main", _targetDir, false);
                if (pushRes.ExitCode == 0)
                {
                    txtCommitMsg.Clear();
                    AppendLog("完了: GitHub上の VBA_Template を最新版に更新しました！");
                }
                await RefreshHistoryListAsync();
            }
            finally
            {
                LoadDirectoryContext();
            }
        }

        private async Task RefreshHistoryListAsync()
        {
            lstHistory.Items.Clear();
            var res = await RunCommandAsync("git", "log -n 30 --date=short --pretty=format:\"%h  %ad  %d %s\"", _targetDir, true);
            if (res.ExitCode == 0 && !string.IsNullOrWhiteSpace(res.Output))
            {
                var lines = res.Output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    lstHistory.Items.Add(line);
                }
            }
        }

        // =====================================================================
        // Context Menu Registry (HKCU - No Admin Required)
        // =====================================================================
        private bool IsContextMenuRegistered()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RegKeyBg))
            {
                return key != null;
            }
        }

        private void UpdateRegButtonText()
        {
            if (IsContextMenuRegistered())
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
            try
            {
                if (IsContextMenuRegistered())
                {
                    Registry.CurrentUser.DeleteSubKeyTree(RegKeyBg, false);
                    Registry.CurrentUser.DeleteSubKeyTree(RegKeyDir, false);
                    AppendLog("右クリックメニューから「VBA Git 管理ツール」を解除しました。");
                }
                else
                {
                    using (var keyBg = Registry.CurrentUser.CreateSubKey(RegKeyBg))
                    {
                        keyBg.SetValue("", "VBA Git 管理ツールを開く");
                        keyBg.SetValue("Icon", _exePath);
                        using (var cmdBg = keyBg.CreateSubKey("command"))
                        {
                            cmdBg.SetValue("", string.Format("\"{0}\" \"%V\"", _exePath));
                        }
                    }
                    using (var keyDir = Registry.CurrentUser.CreateSubKey(RegKeyDir))
                    {
                        keyDir.SetValue("", "VBA Git 管理ツールを開く");
                        keyDir.SetValue("Icon", _exePath);
                        using (var cmdDir = keyDir.CreateSubKey("command"))
                        {
                            cmdDir.SetValue("", string.Format("\"{0}\" \"%1\"", _exePath));
                        }
                    }
                    AppendLog("右クリックメニューに「VBA Git 管理ツールを開く」を登録しました！（管理者権限不要・ユーザー領域）");
                }
                UpdateRegButtonText();
            }
            catch (Exception ex)
            {
                MessageBox.Show("レジストリ操作中にエラーが発生しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        // Process Runner & Helpers
        // =====================================================================
        private class CmdResult
        {
            public int ExitCode;
            public string Output;
        }

        private Task<CmdResult> RunCommandAsync(string fileName, string arguments, string workDir, bool silent)
        {
            return Task.Run(() =>
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    WorkingDirectory = workDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = fileName.Equals("git", StringComparison.OrdinalIgnoreCase) ? Encoding.UTF8 : Encoding.Default,
                    StandardErrorEncoding = fileName.Equals("git", StringComparison.OrdinalIgnoreCase) ? Encoding.UTF8 : Encoding.Default
                };

                var sb = new StringBuilder();
                using (var proc = new Process { StartInfo = psi })
                {
                    proc.OutputDataReceived += (s, e) =>
                    {
                        if (e.Data != null)
                        {
                            sb.AppendLine(e.Data);
                            if (!silent) AppendLog("  " + e.Data);
                        }
                    };
                    proc.ErrorDataReceived += (s, e) =>
                    {
                        if (e.Data != null)
                        {
                            sb.AppendLine(e.Data);
                            if (!silent) AppendLog("  " + e.Data);
                        }
                    };

                    proc.Start();
                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();
                    proc.WaitForExit();

                    return new CmdResult { ExitCode = proc.ExitCode, Output = sb.ToString() };
                }
            });
        }

        private void AppendLog(string text)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action<string>(AppendLog), text);
                return;
            }
            txtLog.AppendText(text + Environment.NewLine);
        }

        private void SetAllButtonsEnabled(bool enabled)
        {
            foreach (Control c in this.Controls)
            {
                if (c is GroupBox)
                {
                    foreach (Control child in c.Controls)
                    {
                        if (child is Button) child.Enabled = enabled;
                    }
                }
            }
        }

        private string PromptInput(string title, string prompt, string defaultVal)
        {
            using (Form f = new Form())
            {
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
    // Built-in Help & Manual Window
    // =========================================================================
    public class HelpForm : Form
    {
        public HelpForm()
        {
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
            Button btnClose = new Button { Text = "閉じる", Location = new Point(550, 8), Size = new Size(100, 32), BackColor = Color.White };
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
                "1. Excel側でマクロを編集し、上書き保存（Ctrl + S）する\r\n" +
                "   ※Excelは開いたままの状態で構いません。\r\n\r\n" +
                "2. フォルダ内を右クリック ＞「VBA Git 管理ツールを開く」\r\n\r\n" +
                "3.「コミットメモ」に修正内容を入力する\r\n" +
                "   （空欄のままでも自動で日時メモが入ります）\r\n\r\n" +
                "4. タイミングに合わせてボタンを押す\r\n" +
                "   ・PC内だけ（ローカル）に保存 → [Save (ローカル保存のみ)]\r\n" +
                "   ・GitHubにも送る（バックアップ） → [Save ＆ GitHub送信 (一括)]\r\n\r\n" +
                "------------------------------------------------------------\r\n" +
                "【別のPCで作業を引き継ぐ時（GitHub利用時）】\r\n\r\n" +
                "1. 必ずExcelブックを閉じた状態にする\r\n" +
                "2. フォルダを右クリックしてツールを開き、[Pull (GitHubから最新取得)] を押す\r\n" +
                "3. GitHubから最新コードを取得し、binフォルダ内のExcelへ自動反映されます。";
        }

        private string GetNewProjectGuideText()
        {
            return
                "【新しいVBA案件を始める時のステップ】\r\n\r\n" +
                "■ パターンA：手元のPC内だけで管理する（社外秘・外部サーバー不可の場合）\r\n" +
                "  1. 案件のフォルダを作りたい「親フォルダ」で右クリックして本ツールを開く\r\n" +
                "  2. [新規作成 (ローカル専用)] を押し、フォルダ名を入力する\r\n" +
                "  3. 新しいフォルダが作成され、自動的に切り替わります\r\n" +
                "  4. [初期設定 (.gitignore)] を押し、[標準除外セット]を追加して確定します\r\n" +
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
                "■ パターンA：さっきSaveした直前の状態に戻したい\r\n" +
                "  → [直前のSave状態へ戻す (Latest)] を押します。\r\n" +
                "     最後に Save した時点のコードがExcelブックへ書き戻されます。\r\n\r\n" +
                "■ パターンB：昨日や先週など、特定の履歴・タグ時点に戻したい\r\n" +
                "  1. 画面中央の「過去の履歴・タグ一覧」から戻りたい行をクリックして選択します。\r\n" +
                "  2. [選択した履歴へ復元 (ローカル)] または [GitHubから取得] を押します。\r\n" +
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
                "  上記のSaveに加え、GitHubへバックアップ送信（Push）まで一気に行います。\r\n\r\n" +
                "・GitHub送信のみ (Push)\r\n" +
                "  ローカルにSave済みの履歴や、後から付けたタグをGitHubへ送信します。\r\n\r\n" +
                "・新規作成 / 新規クローン / 初期設定\r\n" +
                "  空のプロジェクト作成、URLからの環境複製、除外ファイル設定を行います。\r\n\r\n" +
                "・テンプレート更新 (VBA_Template専用)\r\n" +
                "  誤操作防止のため、接続先が「VBA_Template」のフォルダを開いている時のみ\r\n" +
                "  有効化され、それ以外のプロジェクトでは自動的にロックされます。\r\n\r\n" +
                "・Excel使用中警告（安全装置）\r\n" +
                "  Pullや復元時にbin内にロックファイル（~$*.xlsm）がある場合は処理を中断します。";
        }
    }
}