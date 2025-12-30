#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#将文件在一个文件夹中按系统编码复制到另一个文件夹对应的系统编码目录下
"""copy_file_to.py

把「项目概览」里的文件，按“系统编码”匹配后，复制到 pick_file 中对应的系统编码文件夹里。

你的例子：
- 源文件：
  /Users/liaoliming/Desktop/pick/项目概览/项目概览+合同关键页-6661449614937301293-三棵树工程有限公司.pdf
  解析出系统编码：6661449614937301293
- 目标目录（在 pick_file 下找到目录名以系统编码开头的文件夹）：
  /Users/liaoliming/Desktop/pick/pick_file/爱德中创公司/6661449614937301293三棵树工程有限公司52.05万
- 动作：把 PDF 复制到该文件夹内（同名则按策略处理）

用法（CLI）：
  python3 copy_file_to.py \
    --overview-root "/Users/liaoliming/Desktop/pick/项目概览" \
    --pick-root "/Users/liaoliming/Desktop/pick/pick_file" \
    --dry-run \
    --report-csv "/Users/liaoliming/Desktop/pick/overview_copy_report.csv"

可选参数：
  --recursive        递归扫描 overview-root（默认只扫当前目录）
  --regex "..."      用正则从文件名提取系统编码（默认提取连续 16~22 位数字）
  --on-exist skip|overwrite|rename
                    目标文件已存在时：跳过/覆盖/自动改名（默认 rename）
  --prefer-accounting "爱德中创公司"
                    若同一系统编码匹配到多个目录，优先选择该核算主体下面的

GUI：
- 默认不带参数直接启动 GUI：
  python3 "copy_file_to copy.py"
- 或显式启动：
  python3 "copy_file_to copy.py" --gui

特殊规则（支付凭证金额更新）：
- 若源文件名以“支付凭证”开头，且末尾存在“-xx万/xx万元”
- 目标目录中若存在同类支付凭证（仅金额不同）
- 则直接覆盖替换旧文件（不受 on-exist 策略影响）
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import List, Optional, Tuple


def prompt(text: str, default: str = "") -> str:
    if default:
        v = input(f"{text} [{default}]: ").strip()
        return v if v else default
    else:
        return input(f"{text}: ").strip()


DEFAULT_CODE_RE = r"(\d{16,22})"  # 覆盖你给的 19 位编码，也兼容 16~22 位


@dataclass
class CopyResult:
    idx: int
    filename: str
    syscode: str
    src_path: str
    dst_dir: str
    dst_path: str
    status: str  # OK / MISS / FAIL / SKIP
    note: str


def extract_syscode(name: str, pattern: re.Pattern) -> Optional[str]:
    m = pattern.search(name)
    return m.group(1) if m else None


def iter_overview_files(root: Path, recursive: bool) -> List[Path]:
    if not root.exists():
        raise FileNotFoundError(f"overview-root 不存在: {root}")

    if recursive:
        files = [p for p in root.rglob("*") if p.is_file()]
    else:
        files = [p for p in root.iterdir() if p.is_file()]

    # 只处理常见文件（你主要是 pdf；也允许 docx/xlsx/png/jpg 等）
    allow = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".png", ".jpg", ".jpeg"}
    return [p for p in files if p.suffix.lower() in allow]


def find_dest_dirs(pick_root: Path, syscode: str) -> List[Path]:
    """在 pick_root 下找所有“目录名以 syscode 开头”的目录。"""
    if not pick_root.exists():
        raise FileNotFoundError(f"pick-root 不存在: {pick_root}")

    matches: List[Path] = []
    # pick_root/核算主体/系统编码...
    for accounting_dir in pick_root.iterdir():
        if not accounting_dir.is_dir():
            continue
        for p in accounting_dir.iterdir():
            if p.is_dir() and p.name.startswith(syscode):
                matches.append(p)

    # 兜底：如果没有找到，递归搜一遍（防止你 pick_file 下面层级不一致）
    if not matches:
        for p in pick_root.rglob("*"):
            if p.is_dir() and p.name.startswith(syscode):
                matches.append(p)

    return matches


def choose_best_dir(candidates: List[Path], prefer_accounting: str = "") -> Optional[Path]:
    if not candidates:
        return None

    if prefer_accounting:
        preferred = [p for p in candidates if prefer_accounting in str(p.parents[0])]
        if preferred:
            candidates = preferred

    # 选“最短路径 + 最短名字”，通常最像规范目录
    candidates.sort(key=lambda p: (len(str(p)), len(p.name), str(p)))
    return candidates[0]


def ensure_unique_path(dst_dir: Path, src_name: str) -> Path:
    """自动改名：foo.pdf -> foo(1).pdf"""
    base = Path(src_name).stem
    ext = Path(src_name).suffix
    candidate = dst_dir / (base + ext)
    if not candidate.exists():
        return candidate

    for i in range(1, 1000):
        candidate = dst_dir / f"{base}({i}){ext}"
        if not candidate.exists():
            return candidate

    # 极端情况
    raise RuntimeError(f"无法生成不重名文件名: {src_name}")


# ===== 支付凭证同类替换（仅金额不同则覆盖旧文件）相关辅助函数 =====
def _voucher_base_stem(name: str) -> Optional[str]:
    """若为“支付凭证”文件名，返回去掉尾部金额后的 stem，用于同类凭证匹配。

    例：
      支付凭证十一科技-666...-广东恒瀚建筑材料有限公司-23.66万元.pdf
      -> 支付凭证十一科技-666...-广东恒瀚建筑材料有限公司

    只处理以“支付凭证”开头且末尾像“xx万/xx万元”的文件名。
    """
    p = Path(name)
    stem = p.stem
    if not stem.startswith("支付凭证"):
        return None

    # 去掉末尾金额：-30万 / -30万元 / -23.66万 / -23.66万元
    new_stem = re.sub(r"-(\d+(?:\.\d+)?)万(?:元)?$", "", stem)
    if new_stem == stem:
        return None
    return new_stem


def _find_existing_voucher(dst_dir: Path, src_name: str) -> Optional[Path]:
    """在目标目录中查找“同类支付凭证”（仅金额不同）的已存在文件路径。"""
    base = _voucher_base_stem(src_name)
    if not base:
        return None

    matches: List[Path] = []
    try:
        for p in dst_dir.iterdir():
            if not p.is_file():
                continue
            b = _voucher_base_stem(p.name)
            if b and b == base:
                matches.append(p)
    except FileNotFoundError:
        return None

    if not matches:
        return None

    # 多个匹配时，选最短路径/名字（尽量稳定）
    matches.sort(key=lambda p: (len(str(p)), len(p.name), str(p)))
    return matches[0]


def copy_one(src: Path, dst_dir: Path, on_exist: str, dry_run: bool) -> Tuple[str, str, str]:
    """返回 status, dst_path, note"""
    dst_dir.mkdir(parents=True, exist_ok=True)

    if on_exist not in {"skip", "overwrite", "rename"}:
        raise ValueError("--on-exist 只能是 skip/overwrite/rename")

    dst_path = dst_dir / src.name

    # 特殊规则：支付凭证若“同类凭证（仅金额不同）”已存在，则删除旧文件，并以新文件名写入（金额更新）
    existing_voucher = _find_existing_voucher(dst_dir, src.name)
    if existing_voucher is not None:
        new_dst = dst_dir / src.name  # 使用新文件名（包含最新金额）

        # 若新文件名已存在（同名同金额的支付凭证已在目标目录），则跳过
        if new_dst.exists():
            if dry_run:
                return "SKIP", str(new_dst), "dry-run（同名支付凭证已存在，跳过）"
            return "SKIP", str(new_dst), "同名支付凭证已存在，跳过"

        if dry_run:
            note = f"dry-run（支付凭证金额更新：删除旧文件并写入新文件名）旧={existing_voucher.name} 新={new_dst.name}"
            return "OK", str(new_dst), note

        try:
            # 先删旧文件（若旧文件就是同名，则跳过删除）
            if existing_voucher.resolve() != new_dst.resolve() and existing_voucher.exists():
                existing_voucher.unlink()

            # 写入新文件名（若同名已存在则覆盖）
            shutil.copy2(src, new_dst)
            note = f"支付凭证金额更新：已替换 旧={existing_voucher.name} 新={new_dst.name}"
            return "OK", str(new_dst), note
        except Exception as e:
            return "FAIL", str(new_dst), f"支付凭证替换失败: {e}"

    if dst_path.exists():
        if on_exist == "skip":
            return "SKIP", str(dst_path), "目标已存在，跳过"
        if on_exist == "rename":
            dst_path = ensure_unique_path(dst_dir, src.name)
        # overwrite：继续用原名覆盖

    if dry_run:
        return "OK", str(dst_path), "dry-run"

    try:
        shutil.copy2(src, dst_path)
        return "OK", str(dst_path), "copied"
    except Exception as e:
        return "FAIL", str(dst_path), f"复制失败: {e}"


def run_gui() -> None:
    """Tkinter 图形界面：选择目录/参数并执行复制。"""

    root = tk.Tk()
    root.title("file -> pick_file 复制工具")
    root.geometry("980x680")

    # ---- Variables ----
    v_overview = tk.StringVar(value="")
    v_pick = tk.StringVar(value="")
    v_report = tk.StringVar(value="")
    v_recursive = tk.BooleanVar(value=False)
    v_dry = tk.BooleanVar(value=True)
    v_regex = tk.StringVar(value=DEFAULT_CODE_RE)
    v_prefer = tk.StringVar(value="")
    v_onexist = tk.StringVar(value="rename")

    # ---- Layout helpers ----
    def row(parent, r: int, label: str, var: tk.StringVar, browse_cmd=None):
        tk.Label(parent, text=label, anchor="w").grid(row=r, column=0, sticky="w", padx=8, pady=6)
        ent = tk.Entry(parent, textvariable=var)
        ent.grid(row=r, column=1, sticky="we", padx=8, pady=6)
        if browse_cmd:
            tk.Button(parent, text="浏览…", command=browse_cmd).grid(row=r, column=2, sticky="e", padx=8, pady=6)
        return ent

    frm = tk.Frame(root)
    frm.pack(fill="x", padx=10, pady=10)
    frm.columnconfigure(1, weight=1)

    def pick_overview():
        p = filedialog.askdirectory(title="选择文件目录")
        if p:
            v_overview.set(p)

    def pick_pickroot():
        p = filedialog.askdirectory(title="选择放入目录")
        if p:
            v_pick.set(p)

    def pick_report():
        p = filedialog.asksaveasfilename(
            title="选择/输入 报告CSV 输出路径",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if p:
            v_report.set(p)

    row(frm, 0, "选择文件所在目录", v_overview, pick_overview)
    row(frm, 1, "选择需放入文件目录", v_pick, pick_pickroot)
    row(frm, 2, "report-csv（可留空）", v_report, pick_report)

    tk.Label(frm, text="系统编码正则（必须有一个分组）", anchor="w").grid(row=3, column=0, sticky="w", padx=8, pady=6)
    tk.Entry(frm, textvariable=v_regex).grid(row=3, column=1, columnspan=2, sticky="we", padx=8, pady=6)

    tk.Label(frm, text="优先核算主体（可留空）", anchor="w").grid(row=4, column=0, sticky="w", padx=8, pady=6)
    tk.Entry(frm, textvariable=v_prefer).grid(row=4, column=1, columnspan=2, sticky="we", padx=8, pady=6)

    opt = tk.Frame(root)
    opt.pack(fill="x", padx=10)

    tk.Checkbutton(opt, text="递归扫描 overview-root", variable=v_recursive).pack(side="left", padx=6, pady=6)
    tk.Checkbutton(opt, text="dry-run（只预览不复制）", variable=v_dry).pack(side="left", padx=6, pady=6)

    tk.Label(opt, text="目标已存在：").pack(side="left", padx=(20, 6))
    tk.OptionMenu(opt, v_onexist, "skip", "overwrite", "rename").pack(side="left", padx=6)

    # ---- Log area ----
    log_box = ScrolledText(root, height=22)
    log_box.pack(fill="both", expand=True, padx=10, pady=10)

    def append_log(msg: str) -> None:
        log_box.insert("end", msg + "\n")
        log_box.see("end")

    # ---- Run button ----
    btns = tk.Frame(root)
    btns.pack(fill="x", padx=10, pady=(0, 10))

    run_btn = tk.Button(btns, text="开始执行", width=14)
    run_btn.pack(side="left")

    def clear_log():
        log_box.delete("1.0", "end")

    tk.Button(btns, text="清空日志", width=14, command=clear_log).pack(side="left", padx=8)

    status_lbl = tk.Label(btns, text="", anchor="w")
    status_lbl.pack(side="left", padx=10)

    def set_running(running: bool) -> None:
        run_btn.configure(state=("disabled" if running else "normal"))

    def do_run():
        overview = v_overview.get().strip()
        pickr = v_pick.get().strip()
        if not overview or not pickr:
            messagebox.showerror("缺少参数", "请先选择 overview-root 和 pick-root")
            return

        clear_log()
        set_running(True)
        status_lbl.configure(text="运行中…")

        def worker():
            try:
                _, summary = run_copy_job(
                    overview_root=Path(overview).expanduser(),
                    pick_root=Path(pickr).expanduser(),
                    report_csv=v_report.get().strip(),
                    recursive=bool(v_recursive.get()),
                    regex=v_regex.get().strip() or DEFAULT_CODE_RE,
                    prefer_accounting=v_prefer.get().strip(),
                    on_exist=v_onexist.get().strip() or "rename",
                    dry_run=bool(v_dry.get()),
                    log=lambda s: root.after(0, append_log, s),
                )
                root.after(
                    0,
                    status_lbl.configure,
                    {"text": f"完成：成功 {summary['ok']}，未找到 {summary['miss']}，失败 {summary['fail']}，跳过 {summary['skip']}"},
                )
                root.after(
                    0,
                    messagebox.showinfo,
                    "完成",
                    f"总数 {summary['total']}\n成功 {summary['ok']}\n未找到 {summary['miss']}\n失败 {summary['fail']}\n跳过 {summary['skip']}",
                )
            except Exception as e:
                root.after(0, append_log, f"运行失败: {e}")
                root.after(0, status_lbl.configure, {"text": "运行失败"})
                root.after(0, messagebox.showerror, "运行失败", str(e))
            finally:
                root.after(0, set_running, False)

        threading.Thread(target=worker, daemon=True).start()

    run_btn.configure(command=do_run)

    # 允许直接回车执行（在任何 Entry 上）
    def bind_enter(widget):
        widget.bind("<Return>", lambda e: do_run())

    for w in frm.winfo_children():
        if isinstance(w, tk.Entry):
            bind_enter(w)

    root.mainloop()


def run_copy_job(
    overview_root: Path,
    pick_root: Path,
    report_csv: str = "",
    recursive: bool = False,
    regex: str = DEFAULT_CODE_RE,
    prefer_accounting: str = "",
    on_exist: str = "rename",
    dry_run: bool = True,
    log: Optional[callable] = None,
) -> Tuple[List[CopyResult], dict]:
    """核心逻辑：扫描 overview_root，按系统编码匹配 pick_root 下目录并复制。

    Returns:
        results: 每个文件的处理结果
        summary: 统计信息 dict

    log: 可选的日志回调，签名 log(str)
    """

    def _log(msg: str) -> None:
        if log:
            try:
                log(msg)
            except Exception:
                pass

    if not overview_root:
        raise ValueError("overview_root 不能为空")
    if not pick_root:
        raise ValueError("pick_root 不能为空")

    try:
        pattern = re.compile(regex)
    except re.error as e:
        raise ValueError(f"正则不合法: {e}")

    files = iter_overview_files(overview_root, recursive=bool(recursive))
    results: List[CopyResult] = []

    _log("\n===== 开始处理 =====")
    _log(f"overview-root: {overview_root}")
    _log(f"pick-root    : {pick_root}")
    _log(f"recursive    : {recursive}")
    _log(f"dry-run      : {dry_run}")
    _log(f"on-exist     : {on_exist}")
    _log(f"regex        : {regex}")
    _log(f"prefer       : {prefer_accounting or '(空)'}")

    for i, src in enumerate(sorted(files), start=1):
        syscode = extract_syscode(src.name, pattern)
        if not syscode:
            results.append(
                CopyResult(
                    idx=i,
                    filename=src.name,
                    syscode="",
                    src_path=str(src),
                    dst_dir="",
                    dst_path="",
                    status="SKIP",
                    note="文件名未提取到系统编码",
                )
            )
            continue

        candidates = find_dest_dirs(pick_root, syscode)
        dst_dir = choose_best_dir(candidates, prefer_accounting=prefer_accounting)

        if dst_dir is None:
            results.append(
                CopyResult(
                    idx=i,
                    filename=src.name,
                    syscode=syscode,
                    src_path=str(src),
                    dst_dir="",
                    dst_path="",
                    status="MISS",
                    note="pick_file 中未找到以系统编码开头的目录",
                )
            )
            continue

        status, dst_path, note = copy_one(src, dst_dir, on_exist=on_exist, dry_run=bool(dry_run))

        extra = ""
        if len(candidates) > 1:
            extra = f"（同编码匹配到 {len(candidates)} 个目录，已选: {dst_dir}）"

        results.append(
            CopyResult(
                idx=i,
                filename=src.name,
                syscode=syscode,
                src_path=str(src),
                dst_dir=str(dst_dir),
                dst_path=dst_path,
                status=status,
                note=note + extra,
            )
        )

    ok_cnt = sum(1 for r in results if r.status == "OK")
    miss_cnt = sum(1 for r in results if r.status == "MISS")
    fail_cnt = sum(1 for r in results if r.status == "FAIL")
    skip_cnt = sum(1 for r in results if r.status == "SKIP")

    summary = {
        "total": len(results),
        "ok": ok_cnt,
        "miss": miss_cnt,
        "fail": fail_cnt,
        "skip": skip_cnt,
    }

    _log("\n===== 运行摘要 =====")
    _log(f"总文件数: {summary['total']}")
    _log(f"成功: {ok_cnt}  未找到目录: {miss_cnt}  失败: {fail_cnt}  跳过: {skip_cnt}")

    problems = [r for r in results if r.status in {"MISS", "FAIL"}]
    if problems:
        _log("\n===== 问题项（最多 30 条）=====")
        for r in problems[:30]:
            _log(f"#{r.idx} [{r.status}] 编码={r.syscode} 文件={r.filename} 备注={r.note}")

    # 输出报告 CSV（可选）
    if report_csv:
        report = Path(report_csv).expanduser()

        if report.exists() and report.is_dir():
            report = report / "overview_copy_report.csv"
        elif str(report).endswith(("/", "\\")):
            report = Path(str(report).rstrip("/\\")).expanduser() / "overview_copy_report.csv"

        if report.suffix.lower() != ".csv":
            report = report.with_suffix(".csv")

        report.parent.mkdir(parents=True, exist_ok=True)
        with report.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["序号", "文件名", "系统编码", "源文件", "目标目录", "目标文件", "状态", "备注"])
            for r in results:
                w.writerow([r.idx, r.filename, r.syscode, r.src_path, r.dst_dir, r.dst_path, r.status, r.note])
        _log(f"\n报告已输出: {report}")

    return results, summary


def main() -> int:
    ap = argparse.ArgumentParser(description="将项目概览文件按系统编码复制到 pick_file 对应目录")
    ap.add_argument("--overview-root", help="项目概览目录")
    ap.add_argument("--pick-root", help="pick_file 目录")
    ap.add_argument("--recursive", action="store_true", help="递归扫描 overview-root")
    ap.add_argument("--regex", default=DEFAULT_CODE_RE, help="从文件名提取系统编码的正则（必须有一个分组）")
    ap.add_argument("--prefer-accounting", default="", help="多目录匹配时优先的核算主体")
    ap.add_argument("--on-exist", default="rename", choices=["skip", "overwrite", "rename"], help="目标已存在处理方式")
    ap.add_argument("--dry-run", action="store_true", help="只打印不复制")
    ap.add_argument("--report-csv", default="", help="输出报告 CSV")
    ap.add_argument("--gui", action="store_true", help="启动图形界面")

    args = ap.parse_args()

    # 默认行为：不带任何参数时，直接启动图形界面
    if len(sys.argv) == 1 or args.gui:
        run_gui()
        return 0

    print("\n===== 参数填写 =====")

    overview_root_input = args.overview_root or prompt(
        "请输入 文件所在目录 (overview-root)",
    )

    pick_root_input = args.pick_root or prompt(
        "存放分类目录 目录 (pick-root)",
    )

    report_csv_input = args.report_csv or prompt(
        "请输入 报告CSV路径（可留空）",
    )

    dry_run_input = args.dry_run
    if not args.dry_run:
        yn = prompt("是否启用 dry-run（只预览不复制）? y/n", "y")
        dry_run_input = yn.lower().startswith("y")

    regex_input = args.regex or prompt(
        "系统编码正则",
        DEFAULT_CODE_RE
    )

    prefer_accounting_input = args.prefer_accounting or prompt(
        "优先核算主体（可留空）",
        ""
    )

    on_exist_input = args.on_exist or prompt(
        "目标文件已存在时的处理方式 skip / overwrite / rename",
        "rename"
    )

    recursive_input = args.recursive
    if not args.recursive:
        yn = prompt("是否递归扫描 overview-root? y/n", "n")
        recursive_input = yn.lower().startswith("y")

    overview_root = Path(overview_root_input).expanduser()
    pick_root = Path(pick_root_input).expanduser()

    try:
        _, summary = run_copy_job(
            overview_root=overview_root,
            pick_root=pick_root,
            report_csv=report_csv_input,
            recursive=bool(recursive_input),
            regex=regex_input,
            prefer_accounting=prefer_accounting_input,
            on_exist=on_exist_input,
            dry_run=bool(dry_run_input),
            log=print,
        )
    except Exception as e:
        sys.stderr.write(f"运行失败: {e}\n")
        return 2

    return 0 if summary.get("fail", 0) == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())