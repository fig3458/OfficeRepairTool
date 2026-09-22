using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace OfficeRepairTool
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    sealed class MainForm : Form
    {
        static readonly string[] OfficeProcessNames = {
            "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "ONENOTEM",
            "MSPUB", "MSACCESS", "VISIO", "WINPROJ", "GROOVE", "LYNC"
        };

        static readonly string[] RoamingFolders = {
            "Office", "Excel", "Word", "PowerPoint", "Templates"
        };

        static readonly string[] LocalFolders = { "Office", "Excel" };

        readonly Label _status;
        readonly TextBox _log;
        readonly Button _btnKill;
        readonly Button _btnFix;
        readonly PictureBox _pic;

        public MainForm()
        {
            Text = "Office 一键修复工具 — Word 启动失败";
            Width = 740;
            Height = 820;
            MinimumSize = new Size(640, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(12),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            Controls.Add(root);

            var title = new Label
            {
                Text = "用于修复：Word 上次启动失败 / 安全模式提示",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 4),
            };
            root.Controls.Add(title, 0, 0);

            var tip = new Label
            {
                Text = "下方为报错参考图。执行前请先关闭所有 Office 窗口；若关不掉，可用「一键结束 Office 进程」。",
                AutoSize = true,
                MaximumSize = new Size(680, 0),
                ForeColor = Color.FromArgb(68, 68, 68),
                Margin = new Padding(0, 0, 0, 6),
            };
            root.Controls.Add(tip, 0, 1);

            var imgGroup = new GroupBox
            {
                Text = "报错参考（你遇到的窗口）",
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(8),
                Margin = new Padding(0, 0, 0, 6),
            };
            _pic = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.Zoom,
                Width = 660,
                Height = 220,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(238, 238, 238),
            };
            LoadErrorImage();
            imgGroup.Controls.Add(_pic);
            root.Controls.Add(imgGroup, 0, 2);

            var statusPanel = new Panel { Height = 28, Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 6) };
            _status = new Label
            {
                Text = "正在检测 Office 进程…",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(160, 80, 0),
            };
            var btnRefresh = new Button
            {
                Text = "刷新状态",
                Width = 90,
                Dock = DockStyle.Right,
            };
            btnRefresh.Click += (s, e) => RefreshProcessStatus();
            statusPanel.Controls.Add(_status);
            statusPanel.Controls.Add(btnRefresh);
            root.Controls.Add(statusPanel, 0, 3);

            var btnPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4),
            };
            _btnKill = new Button
            {
                Text = "一键结束 Office 进程",
                Width = 180,
                Height = 36,
                Margin = new Padding(0, 0, 8, 0),
            };
            _btnKill.Click += OnKill;
            _btnFix = new Button
            {
                Text = "一键执行修复（重命名注册表 + 缓存）",
                Width = 320,
                Height = 36,
                BackColor = Color.FromArgb(11, 92, 171),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
            };
            _btnFix.FlatAppearance.BorderSize = 0;
            _btnFix.Click += OnFix;
            btnPanel.Controls.Add(_btnKill);
            btnPanel.Controls.Add(_btnFix);
            root.Controls.Add(btnPanel, 0, 4);

            var logGroup = new GroupBox
            {
                Text = "操作日志",
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
            };
            _log = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.White,
            };
            logGroup.Controls.Add(_log);
            root.Controls.Add(logGroup, 0, 5);

            Log("就绪。建议顺序：① 关闭 Office → ②（可选）一键结束进程 → ③ 一键执行修复。");
            Log("修复方式：将相关项重命名为 *.old（不删除），便于出问题后还原。");
            Log("当前：已通过清单申请管理员权限。");
            RefreshProcessStatus();
        }

        void LoadErrorImage()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("word_error.png"))
                {
                    if (stream != null)
                    {
                        _pic.Image = Image.FromStream(stream);
                        return;
                    }
                }
                var beside = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "word_safe_mode_error.png");
                if (File.Exists(beside))
                {
                    _pic.Image = Image.FromFile(beside);
                    return;
                }
            }
            catch { }
            _pic.Height = 60;
            // fallback text via parent — keep empty image area note
        }

        void Log(string msg)
        {
            if (_log.Text.Length > 0) _log.AppendText(Environment.NewLine);
            _log.AppendText(msg);
            _log.SelectionStart = _log.Text.Length;
            _log.ScrollToCaret();
            Application.DoEvents();
        }

        List<ProcessInfo> ListOfficeProcesses()
        {
            var list = new List<ProcessInfo>();
            foreach (var name in OfficeProcessNames)
            {
                Process[] procs;
                try { procs = Process.GetProcessesByName(name); }
                catch { continue; }
                foreach (var p in procs)
                {
                    try
                    {
                        list.Add(new ProcessInfo { Pid = p.Id, Name = p.ProcessName.ToUpperInvariant() + ".EXE" });
                    }
                    catch { }
                    finally { p.Dispose(); }
                }
            }
            return list;
        }

        void RefreshProcessStatus()
        {
            var procs = ListOfficeProcesses();
            if (procs.Count == 0)
            {
                _status.Text = "✓ 未检测到 Office 相关进程，可以执行修复。";
                _status.ForeColor = Color.FromArgb(26, 122, 58);
            }
            else
            {
                var parts = new List<string>();
                foreach (var p in procs) parts.Add(p.Name + "(PID " + p.Pid + ")");
                _status.Text = "⚠ 仍在运行: " + string.Join(", ", parts) + " — 请先关闭，或点「一键结束」。";
                _status.ForeColor = Color.FromArgb(160, 80, 0);
            }
        }

        void OnKill(object sender, EventArgs e)
        {
            var procs = ListOfficeProcesses();
            if (procs.Count == 0)
            {
                MessageBox.Show(this, "当前没有检测到 Office 进程。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshProcessStatus();
                return;
            }
            if (MessageBox.Show(this,
                    "将强制结束 Word / Excel / PowerPoint 等 Office 进程。\n请确认已保存所有文档。\n\n继续吗？",
                    "确认结束进程", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            Log("--- 结束 Office 进程 ---");
            foreach (var name in OfficeProcessNames)
            {
                Process[] running;
                try { running = Process.GetProcessesByName(name); }
                catch { continue; }
                foreach (var p in running)
                {
                    try
                    {
                        var id = p.Id;
                        var n = p.ProcessName.ToUpperInvariant() + ".EXE";
                        p.Kill();
                        if (!p.WaitForExit(3000))
                        {
                            // ignore
                        }
                        Log("已结束进程: " + n + " (PID " + id + ")");
                    }
                    catch (Exception ex)
                    {
                        Log("结束失败: " + name + " — " + ex.Message);
                    }
                    finally { p.Dispose(); }
                }
            }
            System.Threading.Thread.Sleep(500);
            RefreshProcessStatus();
            if (ListOfficeProcesses().Count > 0)
                MessageBox.Show(this, "部分进程未能结束，请手动关闭后重试。", "仍有残留",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show(this, "Office 进程已全部结束。", "完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void OnFix(object sender, EventArgs e)
        {
            if (ListOfficeProcesses().Count > 0)
            {
                if (MessageBox.Show(this,
                        "检测到 Office 进程仍在运行，部分缓存文件夹可能无法重命名。\n\n建议先点「一键结束 Office 进程」。\n\n仍要继续执行修复吗？",
                        "仍有 Office 在运行", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    return;
            }

            if (MessageBox.Show(this,
                    "将执行：\n" +
                    "1. 重命名注册表 HKCU\\Software\\Microsoft\\Office → Office.old\n" +
                    "2. 重命名 %AppData%\\Microsoft 下 Office/Excel/Word/PowerPoint/Templates\n" +
                    "3. 重命名 %LocalAppData%\\Microsoft 下 Office/Excel\n\n" +
                    "均为重命名（加 .old），不会直接删除。\n继续吗？",
                    "确认修复", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            _btnFix.Enabled = false;
            _btnKill.Enabled = false;
            try
            {
                RunRepair();
                MessageBox.Show(this,
                    "修复步骤已执行完毕。\n请重新打开 Word 验证是否还弹出安全模式提示。",
                    "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                _btnFix.Enabled = true;
                _btnKill.Enabled = true;
                RefreshProcessStatus();
            }
        }

        void RunRepair()
        {
            Log("======== 开始修复 ========");
            Log("时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            var running = ListOfficeProcesses();
            if (running.Count > 0)
            {
                var parts = new List<string>();
                foreach (var p in running) parts.Add(p.Name + "(" + p.Pid + ")");
                Log("警告: 仍有 Office 进程在运行: " + string.Join(", ", parts));
            }

            Log("--- 1/3 清理 HKCU 注册表 ---");
            RenameHkcuOfficeKey();

            Log("--- 2/3 清理 Roaming 缓存 ---");
            var roaming = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft");
            foreach (var name in RoamingFolders)
                RenamePath(Path.Combine(roaming, name));

            Log("--- 3/3 清理 Local 缓存 ---");
            var local = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft");
            foreach (var name in LocalFolders)
                RenamePath(Path.Combine(local, name));

            Log("======== 修复完成 ========");
            Log("请重新打开 Word 试一下。若仍异常，可把 *.old 改回原名进行还原。");
        }

        void RenamePath(string path)
        {
            if (!Directory.Exists(path) && !File.Exists(path))
            {
                Log("跳过（不存在）: " + path);
                return;
            }
            var parent = Path.GetDirectoryName(path);
            var name = Path.GetFileName(path);
            var target = Path.Combine(parent, name + ".old");
            if (Directory.Exists(target) || File.Exists(target))
                target = Path.Combine(parent, name + ".old." + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            try
            {
                if (Directory.Exists(path))
                    Directory.Move(path, target);
                else
                    File.Move(path, target);
                Log("已重命名: " + name + " → " + Path.GetFileName(target));
            }
            catch (Exception ex)
            {
                Log("失败: " + path + " — " + ex.Message);
            }
        }

        void RenameHkcuOfficeKey()
        {
            // PowerShell Rename-Item 对注册表更稳妥（.NET 4 无原生 Rename）
            var ps = @"
$ErrorActionPreference = 'Stop'
$path = 'HKCU:\Software\Microsoft\Office'
if (-not (Test-Path $path)) {
    Write-Output 'SKIP: HKCU\Software\Microsoft\Office 不存在'
    exit 0
}
$parent = 'HKCU:\Software\Microsoft'
$newName = 'Office.old'
if (Test-Path (Join-Path $parent $newName)) {
    $newName = 'Office.old.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
}
Rename-Item -Path $path -NewName $newName
Write-Output ('OK: Office -> ' + $newName)
";
            try
            {
                var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(ps));
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand " + encoded,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                };
                using (var p = Process.Start(psi))
                {
                    var stdout = p.StandardOutput.ReadToEnd();
                    var stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit(30000);
                    var outTrim = (stdout ?? "").Trim();
                    var errTrim = (stderr ?? "").Trim();
                    if (p.ExitCode == 0)
                        Log("注册表: " + (string.IsNullOrEmpty(outTrim) ? "完成" : outTrim));
                    else
                        Log("注册表失败: " + (string.IsNullOrEmpty(errTrim) ? outTrim : errTrim));
                }
            }
            catch (Exception ex)
            {
                Log("注册表异常: " + ex.Message);
            }
        }

        sealed class ProcessInfo
        {
            public int Pid;
            public string Name;
        }
    }
}
