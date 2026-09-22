# -*- coding: utf-8 -*-
"""Office / Word 启动失败一键修复工具（Python / tkinter 轻量版）。"""

from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
import sys
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

# ---------------------------------------------------------------------------
# 管理员
# ---------------------------------------------------------------------------

def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def elevate_and_exit() -> None:
    if getattr(sys, "frozen", False):
        executable, args = sys.executable, ""
    else:
        executable = sys.executable
        args = f'"{Path(__file__).resolve()}"'
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", executable, args, None, 1
    )
    if ret <= 32:
        ctypes.windll.user32.MessageBoxW(
            0,
            "未能获取管理员权限。请右键 →「以管理员身份运行」。",
            "无法提权",
            0x10,
        )
    sys.exit(0)


def resource_path(*parts: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base.joinpath(*parts)


# ---------------------------------------------------------------------------
# 进程 / 服务
# ---------------------------------------------------------------------------

# 前台应用
OFFICE_APP_NAMES = {
    "WINWORD.EXE",
    "EXCEL.EXE",
    "POWERPNT.EXE",
    "OUTLOOK.EXE",
    "ONENOTE.EXE",
    "ONENOTEM.EXE",
    "MSPUB.EXE",
    "MSACCESS.EXE",
    "VISIO.EXE",
    "WINPROJ.EXE",
    "GROOVE.EXE",
    "LYNC.EXE",
}

# 后台 / 占用缓存常见进程（不结束它们时 Local\\Office 常无法重命名，强行半改会坏）
OFFICE_BG_NAMES = {
    "OFFICECLICKTORUN.EXE",
    "OFFICEC2RCLIENT.EXE",
    "APPVSHNOTIFY.EXE",
    "MSOSYNC.EXE",
    "MSOUC.EXE",
    "INTEGRATOR.EXE",
    "CSPREFIX.EXE",
}

ALL_OFFICE_NAMES = OFFICE_APP_NAMES | OFFICE_BG_NAMES

ROAMING_FOLDERS = ("Office", "Excel", "Word", "PowerPoint", "Templates")
LOCAL_FOLDERS = ("Office", "Excel")


def _create_no_window() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def list_named_processes(names: set[str]) -> list[tuple[int, str]]:
    found: list[tuple[int, str]] = []
    try:
        import psutil

        for proc in psutil.process_iter(["pid", "name"]):
            name = (proc.info.get("name") or "").upper()
            if name in names:
                found.append((proc.info["pid"], name))
        return found
    except ImportError:
        pass

    try:
        out = subprocess.check_output(
            ["tasklist", "/FO", "CSV", "/NH"],
            creationflags=_create_no_window(),
            text=True,
            encoding="gbk",
            errors="ignore",
        )
    except Exception:
        return found
    for line in out.splitlines():
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) < 2:
            continue
        name, pid_s = parts[0].upper(), parts[1]
        if name in names:
            try:
                found.append((int(pid_s), name))
            except ValueError:
                pass
    return found


def list_office_apps() -> list[tuple[int, str]]:
    return list_named_processes(OFFICE_APP_NAMES)


def list_office_all() -> list[tuple[int, str]]:
    return list_named_processes(ALL_OFFICE_NAMES)


def kill_processes(names: set[str], log) -> list[str]:
    lines: list[str] = []
    try:
        import psutil

        for proc in psutil.process_iter(["pid", "name"]):
            name = (proc.info.get("name") or "").upper()
            if name not in names:
                continue
            pid = proc.info["pid"]
            try:
                p = psutil.Process(pid)
                p.terminate()
                try:
                    p.wait(timeout=3)
                except psutil.TimeoutExpired:
                    p.kill()
                msg = f"已结束进程: {name} (PID {pid})"
                lines.append(msg)
                log(msg)
            except (psutil.NoSuchProcess, psutil.AccessDenied) as exc:
                msg = f"结束失败: {name} (PID {pid}) — {exc}"
                lines.append(msg)
                log(msg)
        return lines
    except ImportError:
        pass

    for pid, name in list_named_processes(names):
        r = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F", "/T"],
            capture_output=True,
            text=True,
            encoding="gbk",
            errors="ignore",
            creationflags=_create_no_window(),
        )
        if r.returncode == 0:
            msg = f"已结束进程: {name} (PID {pid})"
        else:
            msg = f"结束失败: {name} (PID {pid})"
        lines.append(msg)
        log(msg)
    return lines


def stop_clicktorun_service(log) -> bool:
    """暂停 Click-to-Run 服务，避免锁住 Local\\Office。"""
    try:
        r = subprocess.run(
            ["sc", "stop", "ClickToRunSvc"],
            capture_output=True,
            text=True,
            encoding="gbk",
            errors="ignore",
            creationflags=_create_no_window(),
        )
        # 1052 = 未启动, 1062 = 未启动；0 = 成功
        out = ((r.stdout or "") + (r.stderr or "")).strip()
        if r.returncode == 0:
            log("已停止服务: ClickToRunSvc")
            time.sleep(1.2)
            return True
        if "1062" in out or "1052" in out or "没有启动" in out or "not been started" in out.lower():
            log("ClickToRunSvc 本来就未运行")
            return True
        log(f"停止 ClickToRunSvc 返回码 {r.returncode}: {out or '(无输出)'}")
        return False
    except Exception as exc:
        log(f"停止 ClickToRunSvc 异常: {exc}")
        return False


def start_clicktorun_service(log) -> None:
    try:
        r = subprocess.run(
            ["sc", "start", "ClickToRunSvc"],
            capture_output=True,
            text=True,
            encoding="gbk",
            errors="ignore",
            creationflags=_create_no_window(),
        )
        out = ((r.stdout or "") + (r.stderr or "")).strip()
        if r.returncode == 0:
            log("已重新启动服务: ClickToRunSvc")
        else:
            # 1056 = 已在运行
            log(f"启动 ClickToRunSvc: 返回码 {r.returncode} {out}")
    except Exception as exc:
        log(f"启动 ClickToRunSvc 异常: {exc}")


def is_in_use_error(exc: BaseException) -> bool:
    text = str(exc).lower()
    keys = (
        "being used",
        "access is denied",
        "拒绝访问",
        "正在使用",
        "另一个程序",
        "cannot access",
        "winerror 5",
        "winerror 32",
        "errno 13",
        "errno 32",
    )
    if any(k in text for k in keys):
        return True
    if isinstance(exc, PermissionError):
        return True
    if isinstance(exc, OSError) and getattr(exc, "winerror", None) in (5, 32):
        return True
    return False


# ---------------------------------------------------------------------------
# 重命名 / 还原
# ---------------------------------------------------------------------------

def unique_old_name(path: Path) -> Path:
    candidate = path.with_name(path.name + ".old")
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.name}.old.{stamp}")


def rename_path(path: Path, log) -> str:
    """返回 ok / skip / fail。"""
    if not path.exists():
        log(f"跳过（不存在）: {path}")
        return "skip"
    target = unique_old_name(path)
    try:
        path.rename(target)
        log(f"已重命名: {path.name} → {target.name}")
        return "ok"
    except Exception as exc1:
        try:
            shutil.move(str(path), str(target))
            log(f"已移动重命名: {path.name} → {target.name}")
            return "ok"
        except Exception as exc2:
            log(f"失败: {path}")
            log(f"  原因: {exc1} / {exc2}")
            if is_in_use_error(exc1) or is_in_use_error(exc2):
                procs = list_office_all()
                if procs:
                    detail = ", ".join(f"{n}(PID {p})" for p, n in procs)
                    log(f"  疑似占用进程仍在运行: {detail}")
                else:
                    log("  典型原因: 文件/目录正被占用（Office 后台或资源管理器预览）。")
            return "fail"


def rename_hkcu_office_key(log) -> str:
    ps = r"""
$ErrorActionPreference = 'Stop'
$path = 'HKCU:\Software\Microsoft\Office'
if (-not (Test-Path $path)) {
    Write-Output 'SKIP'
    exit 0
}
$parent = 'HKCU:\Software\Microsoft'
$newName = 'Office.old'
if (Test-Path (Join-Path $parent $newName)) {
    $newName = 'Office.old.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
}
Rename-Item -Path $path -NewName $newName
Write-Output ("OK:" + $newName)
"""
    try:
        r = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=_create_no_window(),
        )
        out = (r.stdout or "").strip()
        err = (r.stderr or "").strip()
        if r.returncode == 0:
            if out.startswith("SKIP"):
                log("注册表: HKCU\\Software\\Microsoft\\Office 不存在，跳过")
                return "skip"
            log(f"注册表: 已重命名 Office → {out.replace('OK:', '') or 'Office.old'}")
            return "ok"
        log(f"注册表失败: {err or out or f'exit {r.returncode}'}")
        return "fail"
    except Exception as exc:
        log(f"注册表异常: {exc}")
        return "fail"


def find_old_candidates(parent: Path, base: str) -> list[Path]:
    """找到 base.old 或 base.old.*（取最新）。"""
    if not parent.exists():
        return []
    exact = parent / f"{base}.old"
    timed = sorted(parent.glob(f"{base}.old.*"), key=lambda p: p.stat().st_mtime, reverse=True)
    result: list[Path] = []
    if exact.exists():
        result.append(exact)
    result.extend(timed)
    return result


def restore_path(parent: Path, base: str, log) -> str:
    current = parent / base
    if current.exists():
        log(f"还原跳过（当前已存在）: {current}")
        return "skip"
    candidates = find_old_candidates(parent, base)
    if not candidates:
        log(f"还原跳过（无备份）: {base}.old")
        return "skip"
    src = candidates[0]
    try:
        src.rename(current)
        log(f"已还原: {src.name} → {base}")
        return "ok"
    except Exception as exc:
        log(f"还原失败: {src} → {base} — {exc}")
        return "fail"


def restore_hkcu_office_key(log) -> str:
    ps = r"""
$ErrorActionPreference = 'Stop'
$office = 'HKCU:\Software\Microsoft\Office'
if (Test-Path $office) {
    Write-Output 'SKIP_EXISTS'
    exit 0
}
$parent = 'HKCU:\Software\Microsoft'
$candidates = @()
if (Test-Path (Join-Path $parent 'Office.old')) { $candidates += 'Office.old' }
Get-ChildItem $parent -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'Office.old.*' } |
    Sort-Object Name -Descending | ForEach-Object { $candidates += $_.PSChildName }
if ($candidates.Count -eq 0) {
    Write-Output 'SKIP_NONE'
    exit 0
}
$src = $candidates[0]
Rename-Item -Path (Join-Path $parent $src) -NewName 'Office'
Write-Output ("OK:" + $src)
"""
    try:
        r = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=_create_no_window(),
        )
        out = (r.stdout or "").strip()
        err = (r.stderr or "").strip()
        if r.returncode == 0:
            if out.startswith("SKIP_EXISTS"):
                log("注册表还原跳过：Office 键已存在")
                return "skip"
            if out.startswith("SKIP_NONE"):
                log("注册表还原跳过：没有 Office.old")
                return "skip"
            log(f"注册表已还原: {out.replace('OK:', '')} → Office")
            return "ok"
        log(f"注册表还原失败: {err or out}")
        return "fail"
    except Exception as exc:
        log(f"注册表还原异常: {exc}")
        return "fail"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Office 一键修复工具 — Word 启动失败")
        self.geometry("720x820")
        self.minsize(640, 720)
        self.configure(bg="#f5f5f5")
        self._photo = None
        self._build_ui()
        self.after(200, self.refresh_process_status)

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 4}

        ttk.Label(
            self,
            text="用于修复：Word 上次启动失败 / 安全模式提示",
            font=("Microsoft YaHei UI", 12, "bold"),
        ).pack(anchor="w", **pad)

        ttk.Label(
            self,
            text="执行前请先关闭 Office。关不掉可用「结束进程」。"
            "若上次用旧版 C# 工具修坏了，可点「还原备份」。",
            wraplength=680,
            font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", padx=12, pady=(0, 4))

        img_frame = ttk.LabelFrame(self, text="报错参考（你遇到的窗口）")
        img_frame.pack(fill="x", padx=12, pady=4)
        img_path = resource_path("assets", "word_safe_mode_error.png")
        try:
            self._photo = tk.PhotoImage(file=str(img_path))
            w = self._photo.width()
            if w > 660:
                factor = max(1, int((w + 659) / 660))
                self._photo = self._photo.subsample(factor, factor)
            ttk.Label(img_frame, image=self._photo).pack(padx=8, pady=8)
        except Exception as exc:
            ttk.Label(img_frame, text=f"（无法加载参考图: {exc}）").pack(padx=8, pady=8)

        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=12, pady=4)
        self.status_var = tk.StringVar(value="正在检测 Office 进程…")
        ttk.Label(
            status_frame, textvariable=self.status_var, font=("Microsoft YaHei UI", 9)
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(status_frame, text="刷新状态", command=self.refresh_process_status).pack(
            side="right"
        )

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=12, pady=8)

        self.btn_kill = ttk.Button(
            btn_frame, text="一键结束 Office 进程", command=self.on_kill
        )
        self.btn_kill.pack(side="left", padx=(0, 8))

        self.btn_fix = ttk.Button(
            btn_frame, text="一键执行修复", command=self.on_fix
        )
        self.btn_fix.pack(side="left", padx=(0, 8))

        self.btn_restore = ttk.Button(
            btn_frame, text="还原备份 (*.old)", command=self.on_restore
        )
        self.btn_restore.pack(side="left")

        admin_text = "当前：已获得管理员权限" if is_admin() else "当前：非管理员"
        ttk.Label(self, text=admin_text, font=("Microsoft YaHei UI", 8)).pack(
            anchor="w", padx=12
        )

        log_frame = ttk.LabelFrame(self, text="操作日志")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.log_box = scrolledtext.ScrolledText(
            log_frame, height=14, font=("Consolas", 9), wrap=tk.WORD, state="disabled"
        )
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

        self.log("就绪。建议：① 关闭 Office → ② 结束进程 → ③ 执行修复。")
        self.log("修复用重命名 *.old，不删除。修坏了可用「还原备份」。")

    def log(self, msg: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def refresh_process_status(self) -> None:
        apps = list_office_apps()
        bg = [x for x in list_office_all() if x not in apps]
        if not apps and not bg:
            self.status_var.set("✓ 未检测到 Office 相关进程，可以执行修复。")
            return
        parts = []
        if apps:
            parts.append("前台: " + ", ".join(f"{n}({p})" for p, n in apps))
        if bg:
            parts.append("后台: " + ", ".join(f"{n}({p})" for p, n in bg))
        self.status_var.set("⚠ " + " | ".join(parts))

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.btn_kill.configure(state=state)
        self.btn_fix.configure(state=state)
        self.btn_restore.configure(state=state)

    def on_kill(self) -> None:
        procs = list_office_all()
        if not procs:
            messagebox.showinfo("提示", "当前没有检测到 Office 相关进程。")
            self.refresh_process_status()
            return
        detail = "\n".join(f"  • {n} (PID {p})" for p, n in procs)
        if not messagebox.askyesno(
            "确认结束进程",
            "将结束以下 Office 相关进程（含后台）：\n"
            f"{detail}\n\n请确认文档已保存。继续吗？",
        ):
            return
        self.log("--- 结束 Office 进程 ---")
        kill_processes(ALL_OFFICE_NAMES, self.log)
        stop_clicktorun_service(self.log)
        time.sleep(0.8)
        self.refresh_process_status()
        leftover = list_office_all()
        if leftover:
            detail = ", ".join(f"{n}({p})" for p, n in leftover)
            messagebox.showwarning(
                "仍有残留",
                f"以下进程未能结束，修复时可能提示「正在使用」：\n{detail}\n\n"
                "请手动关闭后重试，或在任务管理器中结束。",
            )
        else:
            messagebox.showinfo("完成", "Office 相关进程已结束（Click-to-Run 服务已尝试停止）。")

    def _prompt_in_use(self, what: str) -> bool:
        """占用失败时提示。返回是否重试该项。"""
        procs = list_office_all()
        if procs:
            detail = "\n".join(f"  • {n} (PID {p})" for p, n in procs)
            msg = (
                f"「{what}」无法重命名：可能正被占用。\n\n"
                f"仍在运行的关联进程：\n{detail}\n\n"
                "是否立即结束这些进程并重试？"
            )
        else:
            msg = (
                f"「{what}」无法重命名（正在使用/拒绝访问）。\n\n"
                "未检测到常见 Office 进程，可能是资源管理器预览、杀毒或"
                "Click-to-Run 服务占用。\n\n"
                "是否再尝试停止 Click-to-Run 并重试一次？"
            )
        return messagebox.askyesno("文件正在使用", msg)

    def on_fix(self) -> None:
        apps = list_office_apps()
        if apps:
            detail = ", ".join(f"{n}({p})" for p, n in apps)
            if not messagebox.askyesno(
                "仍有 Office 在运行",
                f"检测到前台进程仍在运行：\n{detail}\n\n"
                "强烈建议先点「一键结束 Office 进程」。\n\n仍要继续吗？",
            ):
                return

        if not messagebox.askyesno(
            "确认修复",
            "将执行：\n"
            "1. 暂停 Click-to-Run 服务（减少占用）\n"
            "2. 重命名注册表 HKCU\\Software\\Microsoft\\Office → Office.old\n"
            "3. 重命名 Roaming 下 Office/Excel/Word/PowerPoint/Templates\n"
            "4. 重命名 Local 下 Office/Excel\n\n"
            "均为 *.old 备份，不删除。继续吗？",
        ):
            return

        self._set_busy(True)
        try:
            ok, fail = self._run_repair()
            self.refresh_process_status()
            if fail:
                messagebox.showwarning(
                    "修复未完全成功",
                    f"成功 {ok} 项，失败 {fail} 项。\n\n"
                    "失败项多半是「正在使用」。请查看日志中的占用进程提示，\n"
                    "结束后再点一次「一键执行修复」，或用「还原备份」回退。\n\n"
                    "在全部成功前，不建议强行打开 Word。",
                )
            else:
                messagebox.showinfo(
                    "完成",
                    f"修复完成（成功 {ok} 项）。\n请重新打开 Word 验证。",
                )
        finally:
            self._set_busy(False)

    def _run_repair(self) -> tuple[int, int]:
        self.log("======== 开始修复 ========")
        self.log(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        ok = fail = 0

        # 先尽量清进程 + 停服务，避免 Local\\Office 半改坏启动
        leftover = list_office_all()
        if leftover:
            detail = ", ".join(f"{n}({p})" for p, n in leftover)
            self.log(f"警告: 修复前仍有关联进程: {detail}")
        self.log("--- 准备: 结束占用进程 / 停止 ClickToRunSvc ---")
        kill_processes(ALL_OFFICE_NAMES, self.log)
        stop_clicktorun_service(self.log)
        time.sleep(0.8)

        self.log("--- 1/3 清理 HKCU 注册表 ---")
        result = rename_hkcu_office_key(self.log)
        if result == "ok":
            ok += 1
        elif result == "fail":
            fail += 1

        self.log("--- 2/3 清理 Roaming 缓存 ---")
        roaming = Path(os.environ.get("APPDATA", "")) / "Microsoft"
        for name in ROAMING_FOLDERS:
            result = self._rename_with_retry(roaming / name, f"Roaming\\{name}")
            if result == "ok":
                ok += 1
            elif result == "fail":
                fail += 1

        self.log("--- 3/3 清理 Local 缓存 ---")
        local = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft"
        for name in LOCAL_FOLDERS:
            result = self._rename_with_retry(local / name, f"Local\\{name}")
            if result == "ok":
                ok += 1
            elif result == "fail":
                fail += 1

        self.log("--- 收尾: 重启 ClickToRunSvc ---")
        start_clicktorun_service(self.log)

        self.log("======== 修复结束 ========")
        self.log(f"统计: 成功 {ok}，失败 {fail}")
        if fail:
            self.log("有失败项时 Word 可能仍异常；可「还原备份」或处理占用后重试。")
        else:
            self.log("全部成功。请重新打开 Word 试一下。")
        return ok, fail

    def _rename_with_retry(self, path: Path, label: str) -> str:
        result = rename_path(path, self.log)
        if result != "fail":
            return result
        if not self._prompt_in_use(label):
            return "fail"
        self.log(f"用户选择：结束占用并重试 {label}")
        kill_processes(ALL_OFFICE_NAMES, self.log)
        stop_clicktorun_service(self.log)
        time.sleep(1.0)
        result2 = rename_path(path, self.log)
        if result2 == "fail":
            messagebox.showwarning(
                "仍然失败",
                f"「{label}」重试后仍无法重命名。\n"
                "请关闭资源管理器中对该文件夹的预览/窗口后，再执行一次修复。",
            )
        return result2

    def on_restore(self) -> None:
        if not messagebox.askyesno(
            "确认还原",
            "将尝试把 *.old 备份还原为原名（仅当原名不存在时）：\n"
            "• 注册表 Office.old → Office\n"
            "• Roaming / Local 下对应文件夹\n\n"
            "适用于上次修复后 Word 打不开的情况。继续吗？",
        ):
            return
        self._set_busy(True)
        try:
            self.log("======== 开始还原 ========")
            stop_clicktorun_service(self.log)
            kill_processes(ALL_OFFICE_NAMES, self.log)
            time.sleep(0.8)

            ok = fail = 0
            r = restore_hkcu_office_key(self.log)
            if r == "ok":
                ok += 1
            elif r == "fail":
                fail += 1

            roaming = Path(os.environ.get("APPDATA", "")) / "Microsoft"
            for name in ROAMING_FOLDERS:
                r = restore_path(roaming, name, self.log)
                if r == "ok":
                    ok += 1
                elif r == "fail":
                    fail += 1

            local = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft"
            for name in LOCAL_FOLDERS:
                r = restore_path(local, name, self.log)
                if r == "ok":
                    ok += 1
                elif r == "fail":
                    fail += 1

            start_clicktorun_service(self.log)
            self.log(f"======== 还原结束：成功 {ok}，失败 {fail} ========")
            if fail:
                messagebox.showwarning(
                    "还原未完全成功",
                    f"成功 {ok}，失败 {fail}。请查看日志。",
                )
            else:
                messagebox.showinfo("完成", f"还原完成（成功 {ok} 项）。请再试 Word。")
        finally:
            self._set_busy(False)
            self.refresh_process_status()


def main() -> None:
    if os.name == "nt" and not is_admin():
        elevate_and_exit()
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
